"""IMAP service facade module"""
import imaplib
import email
from email.header import decode_header

from auth import authenticate_oauth2


class OutlookService:
    """Facade to hide imaplib complexity"""
    
    def __init__(self, server, email_address, access_token, mailbox="Inbox"):
        self.server = server
        self.email = email_address
        self.token = access_token
        self.mailbox = mailbox
        self.connection = None
    
    def connect(self):
        """Establish connection and authenticate"""
        print(f"[*] Connecting to {self.server}...")
        self.connection = imaplib.IMAP4_SSL(self.server)
        print("[OK] Connection established.")
        
        print(f"[*] Authenticating with OAuth2 for user: {self.email}")
        authenticate_oauth2(self.connection, self.email, self.token)
        print("[OK] OAuth2 login successful.")
        
        # Select mailbox (ReadOnly=False so we can delete later)
        self.connection.select(self.mailbox, readonly=False)
    
    def _is_connection_alive(self):
        """Check if the connection is still alive"""
        try:
            self.connection.noop()
            return True
        except (OSError, imaplib.IMAP4.error, AttributeError):
            return False
    
    def _reconnect(self):
        """Reconnect to the server"""
        print("[*] Connection lost, reconnecting...")
        try:
            if self.connection:
                try:
                    self.connection.close()
                except (OSError, imaplib.IMAP4.error, AttributeError):
                    pass
                try:
                    self.connection.logout()
                except (OSError, imaplib.IMAP4.error, AttributeError):
                    pass
        except (OSError, AttributeError):
            pass
        
        self.connect()
    
    def _try_decode_with_encoding(self, val, encoding):
        """Try to decode bytes with a specific encoding"""
        try:
            return val.decode(encoding, errors='ignore')
        except (UnicodeDecodeError, LookupError):
            return None
    
    def _decode_bytes_with_fallbacks(self, val, encoding):
        """Decode bytes using encoding or fallback to safe encodings"""
        # Try original encoding if valid
        if encoding and encoding.lower() not in ['unknown-8bit', 'unknown']:
            result = self._try_decode_with_encoding(val, encoding)
            if result is not None:
                return result
        
        # Fallback to safe encodings
        for fallback_encoding in ['utf-8', 'latin-1', 'iso-8859-1']:
            result = self._try_decode_with_encoding(val, fallback_encoding)
            if result is not None:
                return result
        
        # Last resort: decode with errors='ignore'
        return val.decode('utf-8', errors='ignore')
    
    def _decode_header_safely(self, header_val):
        """Helper function to decode headers safely with encoding fallbacks"""
        if not header_val:
            return ""
        try:
            decoded_list = decode_header(header_val)
            val, encoding = decoded_list[0]
            if isinstance(val, bytes):
                return self._decode_bytes_with_fallbacks(val, encoding)
            return str(val)
        except Exception:
            # If decode_header fails, return the original as string
            return str(header_val) if header_val else ""
    
    def _fetch_email_subject(self, uid):
        """Fetch subject for a single email UID"""
        try:
            # Use UID FETCH to get persistent identifiers
            typ, msg_data = self.connection.uid('FETCH', uid, '(BODY.PEEK[HEADER.FIELDS (SUBJECT)])')
            if typ != 'OK':
                raise imaplib.IMAP4.error(f"UID FETCH failed: {typ}")
            
            subject = "(No Subject)"
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject_header = msg.get("Subject", "")
                    subject = self._decode_header_safely(subject_header)
            return subject
        except (imaplib.IMAP4.error, OSError):
            # Handle SSL errors, connection errors gracefully
            return "(Error fetching subject)"
    
    def _search_sender_on_server(self, name):
        """Search for emails from a specific sender on the server using UID SEARCH"""
        criteria = f'(FROM "{name}")'
        
        def _uid_search_with_charset(charset=None):
            if not self._is_connection_alive():
                self._reconnect()
            if charset:
                return self.connection.uid('SEARCH', 'CHARSET', charset, criteria)
            # Without charset, rely on server default (usually US-ASCII/UTF-8)
            return self.connection.uid('SEARCH', criteria)
        
        # Try UTF-8 first, fallback to server default if unsupported
        for charset in ('UTF-8', None):
            try:
                typ, data = _uid_search_with_charset(charset)
                if typ != 'OK':
                    raise imaplib.IMAP4.error(f"UID SEARCH failed ({typ})")
                if data and data[0]:
                    return data[0].split()
                return []
            except (imaplib.IMAP4.error, OSError) as e:
                if charset is None:
                    # Already on fallback, log and return empty
                    print(f"  [WARN] Search failed for '{name}': {e}")
                    return []
                # Try next charset option
                continue
        
        return []
    
    def _process_email_ids(self, uids_bytes, sender_name, total_found_uids):
        """Process a list of email UIDs and add them to results"""
        results = []
        
        for uid_bytes in uids_bytes:
            try:
                uid_str = uid_bytes.decode('utf-8')
                
                # Avoid processing the same email twice if it matches multiple keywords
                if uid_str in total_found_uids:
                    continue
                
                total_found_uids.add(uid_str)
                
                # Fetch only the subject for logging purposes
                # If this fails due to SSL error, we still add the email with a default subject
                subject = self._fetch_email_subject(uid_str)
                
                results.append({
                    'id': uid_str,  # Store UID instead of sequence number
                    'sender': sender_name,  # We know the sender because we searched for it
                    'subject': subject
                })
            except (OSError, imaplib.IMAP4.error) as e:
                # If we get a connection error while processing, check connection
                print(f"  [WARN] Connection error processing email from '{sender_name}': {e}")
                # Check if connection is still alive - if not, let exception propagate
                try:
                    self.connection.noop()
                except (OSError, imaplib.IMAP4.error):
                    # Connection is dead, propagate the original exception
                    raise e
                # If connection is alive, continue with next email
                continue
        
        return results
    
    def search_specific_senders(self, sender_names):
        """
        Queries the SERVER directly for specific senders (Server-Side Search).
        This is much faster because it ignores emails from other people.
        
        Instead of downloading all emails and filtering client-side, we ask
        the server to do the filtering, which uses its indexed search.
        
        Args:
            sender_names: List of strings to search for in sender names
                          (e.g., ['Banco Galicia', 'Netflix'])
        
        Returns:
            list: List of dictionaries containing ID, sender, and Subject of found emails
        """
        results = []
        total_found_uids = set()  # Use set to avoid duplicates if partial matches overlap
        
        print(f"[*] Querying server for {len(sender_names)} specific senders...")
        print()
        
        for name in sender_names:
            try:
                uids_bytes = self._search_sender_on_server(name)
                
                if not uids_bytes:
                    # Optimization: If server says "No results", we skip immediately.
                    # We successfully discarded a sender without downloading any headers.
                    print(f"  [SKIP] No emails found from '{name}'")
                    continue
                
                print(f"  [MATCH] Found {len(uids_bytes)} emails from '{name}'")
                
                # Process each found email (using UIDs)
                email_results = self._process_email_ids(uids_bytes, name, total_found_uids)
                results.extend(email_results)
            except (OSError, imaplib.IMAP4.error) as e:
                # Connection error - try to reconnect and continue
                print(f"  [WARN] Connection error while searching '{name}': {e}")
                try:
                    self._reconnect()
                    # Retry the search after reconnection
                    uids_bytes = self._search_sender_on_server(name)
                    if uids_bytes:
                        print(f"  [MATCH] Found {len(uids_bytes)} emails from '{name}' (after reconnect)")
                        email_results = self._process_email_ids(uids_bytes, name, total_found_uids)
                        results.extend(email_results)
                    else:
                        print(f"  [SKIP] No emails found from '{name}' (after reconnect)")
                except Exception as reconnect_err:
                    print(f"  [ERROR] Failed to reconnect: {reconnect_err}")
                    # Continue with next sender
                    continue
        
        return results
    
    def get_message_headers(self):
        """
        DEPRECATED: Use search_specific_senders() for better performance.
        This method downloads all emails and filters client-side.
        
        Returns:
            list: List of dictionaries with email information
        """
        print("[*] Retrieving email list...")
        _, messages = self.connection.search(None, 'ALL')
        
        email_ids_bytes = messages[0].split()
        if not email_ids_bytes:
            print("[INFO] Inbox is empty.")
            return []
        
        total_emails = len(email_ids_bytes)
        print(f"[*] Analyzing {total_emails} emails (Headers only)...")
        
        results = []
        for e_id in email_ids_bytes:
            # CRITICAL OPTIMIZATION: Fetch only headers, not full email
            # This downloads only ~200 bytes per email instead of full body
            # PEEK prevents marking the email as "Read"
            _, msg_data = self.connection.fetch(e_id, '(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])')
            
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    # Parse the bytes into an email object (containing only headers)
                    msg = email.message_from_bytes(response_part[1])
                    
                    # Extract and decode headers
                    from_header = msg.get("From", "")
                    subject_header = msg.get("Subject", "(No Subject)")
                    
                    sender_name = self._decode_header_safely(from_header)
                    subject = self._decode_header_safely(subject_header)
                    
                    # Clean up "Name <email>" format to just "Name"
                    if '<' in sender_name:
                        sender_name = sender_name.split('<')[0].strip().strip('"')
                    
                    results.append({
                        'id': e_id.decode('utf-8'),
                        'sender': sender_name,
                        'subject': subject
                    })
        
        return results
    
    def _delete_single_uid(self, uid):
        """Flag a single UID for deletion"""
        if not self._is_connection_alive():
            self._reconnect()
        self.connection.uid('STORE', uid, '+FLAGS', '\\Deleted')
    
    def _retry_delete_single_uid(self, uid):
        """Retry deleting a single UID after reconnection"""
        self._reconnect()
        self.connection.uid('STORE', uid, '+FLAGS', '\\Deleted')
    
    def _process_deletion_batch(self, batch_uids, batch_start):
        """Process batch by iterating each UID (avoids invalid message-set errors)"""
        for offset, uid in enumerate(batch_uids):
            try:
                self._delete_single_uid(uid)
            except (OSError, imaplib.IMAP4.error) as e:
                print(f"  [WARN] Error deleting UID {uid}: {e}")
                try:
                    self._retry_delete_single_uid(uid)
                except Exception as retry_err:
                    print(f"  [ERROR] Failed to delete UID {uid} after reconnect: {retry_err}")
                    continue
        print(f"    [*] Processed batch {batch_start} to {batch_start + len(batch_uids)}")
    
    def _expunge_with_retry(self):
        """Execute expunge with reconnection retry"""
        try:
            if not self._is_connection_alive():
                self._reconnect()
            self.connection.expunge()
            print("[OK] Cleanup completed successfully.")
        except (OSError, imaplib.IMAP4.error) as e:
            print(f"[WARN] Error during expunge: {e}")
            # Try to reconnect and retry expunge
            try:
                self._reconnect()
                self.connection.expunge()
                print("[OK] Cleanup completed successfully (after reconnect).")
            except Exception as expunge_err:
                print(f"[ERROR] Failed to complete expunge: {expunge_err}")
    
    def delete_emails(self, email_ids, batch_size=100):
        """
        Handles the batch deletion logic
        
        Args:
            email_ids: List of email ID strings to delete
            batch_size: Number of emails to process per batch
        """
        if not email_ids:
            print("[INFO] No emails to delete.")
            return
        
        print(f"[*] Moving {len(email_ids)} emails to Deleted folder in batches...")
        
        # Check connection before starting deletion
        if not self._is_connection_alive():
            self._reconnect()
        
        # IMAP commands can fail if the command line is too long,
        # so we process in batches
        # Note: email_ids are now UIDs, which are persistent across reconnections
        for k in range(0, len(email_ids), batch_size):
            batch = email_ids[k:k+batch_size]
            self._process_deletion_batch(batch, k)
        
        # Apply the deletion (expunge)
        self._expunge_with_retry()
    
    def close(self):
        """Close connection and logout"""
        if self.connection:
            try:
                self.connection.close()
            except Exception:
                pass
            try:
                self.connection.logout()
            except Exception:
                pass


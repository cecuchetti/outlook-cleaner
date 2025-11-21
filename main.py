"""Main entry point for Outlook Cleaner"""
from config import load_config, get_config_value
from auth import get_oauth2_token
from filters import SenderNameFilter, EmailFilter
from imap_service import OutlookService


def clean_inbox(user_email, email_filter: EmailFilter, client_id, tenant_id="consumers", 
                access_token=None, force_interactive=True, move_to_deleted=True,
                imap_server="outlook.office365.com", mailbox="Inbox"):
    """
    Clean emails using Strategy and Facade patterns
    
    Args:
        user_email: User's email address
        email_filter: EmailFilter strategy instance to determine which emails to delete
        client_id: OAuth2 client ID
        tenant_id: OAuth2 tenant ID
        access_token: Optional access token (will be fetched if not provided)
        force_interactive: Force interactive OAuth2 login
        move_to_deleted: If True, move emails to deleted. If False, only list them
        imap_server: IMAP server address
        mailbox: Mailbox to process
    """
    try:
        # Get token if not provided
        if not access_token:
            access_token = get_oauth2_token(client_id, tenant_id, user_email, force_interactive)
        
        # Setup Service (Facade) - hides IMAP complexity
        service = OutlookService(
            server=imap_server,
            email_address=user_email,
            access_token=access_token,
            mailbox=mailbox
        )
        
        try:
            service.connect()
            
            # Show what we're searching for
            print(f"[*] {email_filter.get_description()}")
            print()
            
            # OPTIMIZATION: Use server-side search for SenderNameFilter
            # This queries the server directly instead of downloading all emails
            if isinstance(email_filter, SenderNameFilter):
                # Extract the sender names from the filter (original names, not uppercase)
                sender_names = email_filter.get_sender_names_for_search()
                # Use server-side search (much faster for large inboxes)
                emails = service.search_specific_senders(sender_names)
                
                # All returned emails are matches (server already filtered)
                ids_to_delete = [mail['id'] for mail in emails]
                
                # Print matches
                for mail in emails:
                    print(f"  [MATCH] '{mail['sender']}' | Subject: {mail['subject'][:50]}...")
            else:
                # Fallback to client-side filtering for other filter types
                # (e.g., SubjectFilter that can't use server-side search)
                emails = service.get_message_headers()
                
                # Apply the strategy to filter emails
                ids_to_delete = []
                for mail in emails:
                    if email_filter.matches(mail):
                        matched_info = mail.get('matched_name') or mail.get('matched_keyword', 'N/A')
                        print(f"  [MATCH] '{matched_info}' in sender '{mail['sender']}' | Subject: {mail['subject'][:50]}...")
                        ids_to_delete.append(mail['id'])
            
            print(f"\n[*] Summary: Found {len(ids_to_delete)} emails to process.")
            
            # Delete emails if enabled
            if move_to_deleted and ids_to_delete:
                service.delete_emails(ids_to_delete)
            else:
                print("[INFO] Read-only mode or no matches found. No changes made.")
        
        finally:
            service.close()
    
    except Exception as e:
        print(f"[ERROR] {e}")


def main():
    """
    Main function that loads configuration and executes email cleanup
    """
    try:
        # Load configuration
        config = load_config()
        
        # Extract configuration values
        email = get_config_value(config, 'email')
        if not email:
            raise ValueError("[ERROR] 'email' is required in configuration")
        
        sender_names = get_config_value(config, 'cleaning', 'sender_names_to_search', default=[])
        if not sender_names:
            raise ValueError("[ERROR] 'cleaning.sender_names_to_search' is required in configuration")
        
        client_id = get_config_value(config, 'oauth2', 'client_id')
        if not client_id or client_id == "tu-client-id-aqui":
            raise ValueError("[ERROR] 'oauth2.client_id' must be configured in config.json")
        
        tenant_id = get_config_value(config, 'oauth2', 'tenant_id', default='consumers')
        force_interactive = get_config_value(config, 'oauth2', 'force_interactive_login', default=True)
        mover_a_eliminados = get_config_value(config, 'cleaning', 'move_to_deleted', default=True)
        imap_server = get_config_value(config, 'imap', 'server', default='outlook.office365.com')
        mailbox = get_config_value(config, 'imap', 'mailbox', default='Inbox')
        
        # Setup Strategy - inject the specific rule we want to use
        email_filter = SenderNameFilter(sender_names)
        
        # Execute cleanup using Facade and Strategy patterns
        clean_inbox(
            email,
            email_filter,
            client_id=client_id,
            tenant_id=tenant_id,
            force_interactive=force_interactive,
            move_to_deleted=mover_a_eliminados,
            imap_server=imap_server,
            mailbox=mailbox
        )
        
    except FileNotFoundError as e:
        print(e)
        return 1
    except ValueError as e:
        print(e)
        return 1
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())

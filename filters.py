"""Email filtering strategies module (Strategy Pattern)"""
from abc import ABC, abstractmethod


class EmailFilter(ABC):
    """Abstract Base Class for filtering strategies"""
    
    @abstractmethod
    def matches(self, email_data: dict) -> bool:
        """
        Check if an email matches the filter criteria
        
        Args:
            email_data: Dictionary containing email information (sender, subject, etc.)
            
        Returns:
            bool: True if email matches, False otherwise
        """
        pass
    
    @abstractmethod
    def get_description(self) -> str:
        """Return a description of what this filter is looking for"""
        pass


class SenderNameFilter(EmailFilter):
    """Concrete Strategy: Filter by Sender Name"""
    
    def __init__(self, restricted_names):
        # Store original names for server-side search
        self.restricted_names_original = restricted_names
        # Store uppercase versions for client-side matching
        self.restricted_names = [n.upper() for n in restricted_names]
    
    def matches(self, email_data: dict) -> bool:
        sender = email_data.get('sender', '').upper()
        for name in self.restricted_names:
            if name in sender:
                # Find the original name that matched
                for orig_name in self.restricted_names_original:
                    if orig_name.upper() == name:
                        email_data['matched_name'] = orig_name
                        break
                return True
        return False
    
    def get_description(self) -> str:
        return f"Senders containing: {', '.join(self.restricted_names_original)}"
    
    def get_sender_names_for_search(self):
        """Get original sender names for server-side IMAP search"""
        return self.restricted_names_original


class SubjectFilter(EmailFilter):
    """Concrete Strategy: Filter by Subject Keywords (Example of extensibility)"""
    
    def __init__(self, keywords):
        self.keywords = [k.upper() for k in keywords]
    
    def matches(self, email_data: dict) -> bool:
        subject = email_data.get('subject', '').upper()
        for keyword in self.keywords:
            if keyword in subject:
                email_data['matched_keyword'] = keyword
                return True
        return False
    
    def get_description(self) -> str:
        return f"Subjects containing: {', '.join(self.keywords)}"


"""
Email Service - Handles all email operations.

Consolidates email sending logic that was previously scattered
across Team and EmailSender classes.
"""
import os
import time
import json
from dataclasses import dataclass
from typing import Optional
import pythoncom

from logClass.log import Log
from utils.zip_folder import zip_folder


@dataclass
class SenderInfo:
    """Information about the email sender for signatures."""
    full_name: str = "Client Services Department"
    job_position: str = "Thermo Fisher Bot"
    site_address: str = "Av. Del Campo 1550/60, Ciudad Autónoma de Buenos Aires, CP1427"
    phone_number: str = "-"
    email_address: str = "-"


@dataclass
class OrderSummary:
    """Summary of processed orders for email."""
    team_name: str
    ship_date: str
    total_orders: int
    processed_orders: int
    pending_orders: int


class EmailService:
    """
    Service for sending emails via Outlook.
    
    Handles:
    - Team notification emails with order summaries
    - Medical center notification emails
    - Email signature management
    """
    
    CONFIG_FILE = "emailSender/last_sender_data.json"
    TEAM_EMAIL_TEMPLATE = "media/email_to_team.txt"
    
    def __init__(self, log: Log, sender: Optional[SenderInfo] = None):
        """
        Initialize the email service.
        
        Args:
            log: Logger instance.
            sender: Sender information for signatures.
        """
        self._log = log
        self._sender = sender or SenderInfo()
    
    # --- Public Methods ---
    
    def send_team_summary_email(
        self,
        to_email: str,
        order_summary: OrderSummary,
        attachments_folder: str,
        cc_email: Optional[str] = None,
    ) -> bool:
        """
        Send a summary email to the team with processed orders.
        
        Args:
            to_email: Team email address.
            order_summary: Summary of processed orders.
            attachments_folder: Folder containing files to attach.
            cc_email: Optional CC email address.
            
        Returns:
            True if email was sent successfully.
        """
        try:
            # Create ZIP of attachments
            zip_path = self._create_attachments_zip(
                attachments_folder,
                order_summary.team_name,
                order_summary.ship_date,
            )
            
            # Build email content
            subject = f"Shipping orders with dispatch date {order_summary.ship_date} - {order_summary.team_name}"
            body = self._build_team_email_body(order_summary)
            
            # Send email
            self._send_outlook_email(
                to=to_email,
                cc=cc_email,
                subject=subject,
                html_body=body,
                attachment_path=zip_path,
            )
            
            return True
            
        except Exception as e:
            self._log.add_error_log(f"Error sending team email: {e}")
            return False
    
    def set_sender(self, sender: SenderInfo) -> None:
        """Update the sender information."""
        self._sender = sender
    
    def save_sender_config(self) -> None:
        """Save the current sender configuration to file."""
        config = {
            "full_name": self._sender.full_name,
            "job_position": self._sender.job_position,
            "site_address": self._sender.site_address,
            "phone_number": self._sender.phone_number,
            "email_address": self._sender.email_address,
        }
        
        with open(self.CONFIG_FILE, "w") as f:
            json.dump(config, f)
    
    def load_sender_config(self) -> SenderInfo:
        """Load sender configuration from file."""
        try:
            with open(self.CONFIG_FILE, "r") as f:
                config = json.load(f)
                return SenderInfo(**config)
        except (FileNotFoundError, json.JSONDecodeError):
            return SenderInfo()
    
    # --- Private Methods ---
    
    def _create_attachments_zip(
        self,
        folder_path: str,
        team_name: str,
        date: str,
    ) -> str:
        """Create a ZIP file of the attachments folder."""
        parent_folder = os.path.dirname(folder_path)
        zip_name = f"orders_{team_name}_{date}"
        zip_path = os.path.join(parent_folder, zip_name)
        
        zip_folder(folder_path, zip_path)
        
        return f"{zip_path}.zip"
    
    def _build_team_email_body(self, summary: OrderSummary) -> str:
        """Build the HTML email body from template."""
        template = self._load_email_template(self.TEAM_EMAIL_TEMPLATE)
        
        # Replace order variables
        body = template.replace("|VAR_SELECTED_TEAM|", summary.team_name)
        body = body.replace("|VAR_SHIP_DATE|", summary.ship_date)
        body = body.replace("|VAR_TOTAL_AMOUNT_OF_ORDERS|", str(summary.total_orders))
        body = body.replace("|VAR_AMOUNT_OF_ORDERS_PROCESSED|", str(summary.processed_orders))
        body = body.replace("|VAR_AMOUNT_OF_ORDERS_NOT_PROCESSED|", str(summary.pending_orders))
        body = body.replace("|VAR_TMO_LOGO|", os.getcwd() + "\\media\\TMO_logo_email.jpg")
        
        # Replace signature variables
        body = self._apply_signature(body)
        
        return body
    
    def _apply_signature(self, body: str) -> str:
        """Apply sender signature to email body."""
        body = body.replace("|VAR_FULL_NAME|", self._sender.full_name)
        body = body.replace("|VAR_JOB_POSITION|", self._sender.job_position)
        body = body.replace("|VAR_SITE_ADDRESS|", self._sender.site_address)
        body = body.replace("|VAR_PHONE_NUMBER|", self._sender.phone_number)
        body = body.replace("|VAR_EMAIL_ADDRESS|", self._sender.email_address)
        return body
    
    def _load_email_template(self, template_path: str) -> str:
        """Load an email template from file."""
        with open(template_path, "r", encoding="utf-8") as f:
            return f.read()
    
    def _send_outlook_email(
        self,
        to: str,
        subject: str,
        html_body: str,
        cc: Optional[str] = None,
        attachment_path: Optional[str] = None,
    ) -> None:
        """Send an email using Outlook."""
        import win32com.client as win32
        
        pythoncom.CoInitialize()
        
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            
            mail.To = to
            if cc:
                mail.CC = cc
            mail.Subject = subject
            mail.HTMLBody = html_body
            
            if attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            
            time.sleep(1)
            mail.Send()
            
        finally:
            pythoncom.CoUninitialize()

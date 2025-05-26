import json

class EmailSender:
    def __init__(self, full_name: str = "", 
                job_position: str = "", site_address: str = "",
                phone_number: str = "", email_address: str = ""):
        
        self.file = "emailSender/last_sender_data.json"

        if full_name == "":
            self.full_name = "Client Services Department"
            self.job_position = "Thermo Fisher Bot"
            self.site_address = "Av. Del Campo 1550/60, Ciudad Autónoma de Buenos Aires, CP1427"
            self.phone_number = "-"
            self.email_address = "-"
        else:
            self.full_name = full_name
            self.job_position = job_position
            self.site_address = site_address
            self.phone_number = phone_number
            self.email_address = email_address

    def get_last_sender_config(self) -> dict:
        with open(self.file, "r") as f:
            return json.load(f)
        
    def save_last_sender_config(self, full_name: str, job_position: str, site_address: str, phone_number: str, email_address: str) -> None:
        with open(self.file, "w") as f:
            json.dump({
                "full_name": full_name,
                "job_position": job_position,
                "site_address": site_address,
                "phone_number": phone_number,
                "email_address": email_address
            }, f)

    def replace_email_signature(self, email_body: str) -> str:
        email_body = email_body.replace("|VAR_FULL_NAME|", self.full_name)
        email_body = email_body.replace("|VAR_JOB_POSITION|", self.job_position)
        email_body = email_body.replace("|VAR_SITE_ADDRESS|", self.site_address)
        email_body = email_body.replace("|VAR_PHONE_NUMBER|", self.phone_number)
        email_body = email_body.replace("|VAR_EMAIL_ADDRESS|", self.email_address)

        return email_body
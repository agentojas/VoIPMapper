# VoIPMapper
VoIPMapper is a IPDR intelligence tool that maps destination IP addresses to VoIP platforms and messaging services, Social Media Infrastructure...

---

the script scans the Destination Port of each record.
If the port matches known VoIP ports (for example, 3478 or other configured ports), the record is marked for further analysis.The script checks the Destination IP Address against a database of known IP addresses and domains used by various communication platforms.If required, it also performs DNS lookups to identify the organization or service associated with the IP address.

# Based on the IP address, domain, and network patterns, the script determines the most likely platform used for the VoIP communication, such as:  
- Whatsapp
- Instagram
- Telegram
- Signal Messenger
- Zangi
- Discord
- Twitter ( X ) 
- Snapchat
- Wire
- Viber ( Will Added )
- Threema ( Will Added )
  
# Output

- The identified platform is added to the New IPDR Xlsx File
- The script highlights detected VoIP records using different colors for easier review.
- A new Excel file is generated containing Only detected VoIP platform and other investigation details.

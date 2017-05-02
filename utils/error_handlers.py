from .email_sender import send_mail


def handle_500(e):
    send_mail.delay("ppal@auroim.com",
              ['ppal@auroim.com'],
              "Error Ocurred",
              str(e),
              username='ppal@auroim.com',
              password='AuroOct2016')
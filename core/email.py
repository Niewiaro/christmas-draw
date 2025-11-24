def send_email_via_outlook(
    *,
    to: str,
    subject: str,
    body: str = None,
    html_body: str = None,
) -> None:
    """Sends an email via Outlook."""
    import win32com.client as win32

    if body is None and html_body is None:
        raise ValueError("Either body or html_body must be provided.")

    # Setting up Outlook email
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject

    if body:
        mail.Body = body

    if html_body:
        mail.HTMLBody = html_body
    mail.Send()

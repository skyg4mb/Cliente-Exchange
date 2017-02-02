public class MailService
    {
        /*El metodo SendMail envia un correo utilizando la api de exchange 2007, se deben enviar los parametros
         * From: remitente del correo electronico
         * Torecipients. Destinatario del correo electronico
         * Message: Mensaje de correo 
         * Subject: Asundo del correo 
         * attachment: Adjunto del correo, se debe convertir en tipo byte ejemplo: 
         * attachmentName: Nombre que tendra el adjunto en el correo electronico, debe incluir extension
         * user: Usuario con el cual se enviara el correo
         * password: password del usuario 
         * */
        string from;

        public string From
        {
            get { return from; }
            set { from = value; }
        }
        string user;

        public string User
        {
            get { return user; }
            set { user = value; }
        }


        string password;

        public string Password
        {
            get { return password; }
            set { password = value; }
        }
        string torecipients;

        public string Torecipients
        {
            get { return torecipients; }
            set { torecipients = value; }
        }
        string message;

        public string Message
        {
            get { return message; }
            set { message = value; }
        }
        string subject;

        public string Subject
        {
            get { return subject; }
            set { subject = value; }
        }
        byte[] attachment = null;

        public byte[] Attachment
        {
            get { return attachment; }
            set { attachment = value; }
        }
        string attachmentName="";

        public string AttachmentName
        {
            get { return attachmentName; }
            set { attachmentName = value; }
        }

        public MailService(string From, string User, string Password, string ToRecipients, string Message, 
            string Subject, byte[] Attachment, string AttachmentName)
        {
            from = From;
            user = User;
            password = Password;
            torecipients = ToRecipients;
            message = Message;
            subject = Subject;
            attachment = Attachment;
            attachmentName = AttachmentName;
        }

        //envia correo sin adjunto
        public MailService(string From, string User, string Password, string ToRecipients, string Message,
            string Subject)
        {
            from = From;
            user = User;
            password = Password;
            torecipients = ToRecipients;
            message = Message;
            subject = Subject;
            attachment = Attachment;
            attachmentName = AttachmentName;
        }

        public void Sendmail()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.UseDefaultCredentials = true;
            service.AutodiscoverUrl(from, RedirectionUrlValidationCallback);
            service.Credentials = new WebCredentials(user, password);
            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add(torecipients);
            email.Subject = subject;
            if (attachmentName.Length>0)
            {
                email.Attachments.AddFileAttachment(attachmentName, attachment);
            }
            email.Body = new MessageBody(message);
            email.Send();
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }

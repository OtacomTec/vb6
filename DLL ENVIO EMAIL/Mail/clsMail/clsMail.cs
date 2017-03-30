using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net; // importe o namespace .Net
using System.Net.Mail; // importe o namespace .Net.Mail
//usar o InteropServices - CRIA A DLL COMPATIVEL COM VB6
using System.Runtime.InteropServices;

namespace clsMail
{
    //define o tipo de interface da classe - CRIA A DLL COMPATIVEL COM VB6
    [ClassInterface(ClassInterfaceType.AutoDual)]

    //registra um identificar para a classe no registry - CRIA A DLL COMPATIVEL COM VB6
    [ProgId("clsMail.SmtpEnvio")]

    //faz com que todos os métodos e propriedades da classe sejam visiveis - CRIA A DLL COMPATIVEL COM VB6
    [ComVisible(true)]

    

    public class SmtpEnvio
    {
        public string Envia(string Host, string Port, string From, string FromName, string AddAddress, string Body, string Subject, string UserName, string password,bool habilita_SSL) 
        {

            SmtpClient Client_mail = new SmtpClient(Host.ToString(), Convert.ToInt32(Port));
            Client_mail.EnableSsl = habilita_SSL;

            MailAddress remetente = new MailAddress(From.ToString(), FromName.ToString() );
            MailAddress destinatario = new MailAddress(AddAddress.ToString(), FromName.ToString());
            MailMessage mensagem = new MailMessage(From.ToString(), AddAddress.ToString());

            mensagem.Body = Body.ToString();
            mensagem.Subject = Subject.ToString();
            
            NetworkCredential credenciais = new NetworkCredential(UserName.ToString() , password.ToString() , "");

            Client_mail.Credentials = credenciais;

            try
            {
                Client_mail.Send(mensagem);
                return "Enviado com sucesso!";                
                
            }
            catch (Exception e)
            {
                return e.Message.ToString();
            }
            
        }
    }
}

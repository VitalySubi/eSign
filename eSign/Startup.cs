using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using eSign;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            DigitalSigner ds = new DigitalSigner(@"C:\Users\Vitaly\Pictures\iesk.png");
            /// Вытаскивает из хранилища список сертификатов, из этого списка нужно будет отобрать нужный
            /// либо автоматически, по пользователю, либо чтобы пользователь сам выбирал
            ds.GetMyCerts();
            /// Выбор нужного сертификата для подписи документа. В данном случае выбор по серийному номеру
            ds.GetMyESignCert(X509FindType.FindBySerialNumber, "120060808DC9433CF60A124A1B00010060808D");
            /// Создание УЭЦП и подпись
            ds.CreateAndPlaceSignature(@"C:\Users\Vitaly\Desktop\Теория.pdf", @"C:\Users\Vitaly\Desktop\Теория_signed.pdf");

            DigitalVerifier dv = new DigitalVerifier();
            var res = dv.Verify(@"C:\Users\Vitaly\Desktop\Теория_signed.pdf");
        }
    }
}

using System;
using System.Collections.Generic;
using CAPICOM;
using CAdESCOM;
using iTextSharp.text.pdf;
using System.IO;
using Org.BouncyCastle.X509;
using System.Security.Cryptography.X509Certificates;
using iTextSharp.text;
using System.Text.RegularExpressions;
using iTextSharp.text.io;
using System.Security.Cryptography.Pkcs;

namespace eSign
{
    /// <summary>
    /// Класс для подписи документов
    /// </summary>
    public class DigitalSigner
    {
        private readonly string TSAAddress;
        private readonly string LogoPath;

        private X509Certificate2Collection myCertsCollection;
        private X509Certificate2 myCert;

        public X509Certificate2 MyCert { get { return myCert; } }
        public X509Certificate2Collection MyCertsCollection { get { return myCertsCollection; } }

        public int buff = 10000000;
        
        public string reason = "Подписание документа";
        public string location = "Российская Федерация, Иркутская область, г. Иркутск";
        public string fontPath = "c:/Windows/Fonts/arial.ttf";

        public int? signaturePageNum = null;

        /// <summary>
        /// Цвета полоски "Сведения" в формате RGB
        /// </summary>
        public float[] lineColors = { 0.21F, 0.45F, 0.89F };

        /// <summary>
        /// 
        /// </summary>
        /// <param name="logoPath">Путь к лого</param>
        /// <param name="tsaAddress">Сервис для получения штампа времени</param>
        public DigitalSigner(string logoPath, string tsaAddress= "http://qs.cryptopro.ru/tsp/tsp.srf")
        {
            LogoPath = logoPath;
            TSAAddress = tsaAddress;
        }

        /// <summary>
        /// Выбирает из локального хранилища сертификатов все доступные сертификаты пользователя и 
        /// записывает их в поле myCertsCollection. Для доступа используйте MyCertsCollection. 
        /// </summary>
        public void GetMyCerts()
        {
            X509Store store = new X509Store("My", StoreLocation.CurrentUser);
            store.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadOnly);

            myCertsCollection = store.Certificates;
            store.Close();
        }

        /// <summary>
        /// Выбирает из списка локальных сертификатов пользователя один сертификат, который будет использован
        /// для подписания документа. 
        /// </summary>
        /// <param name="findValue">Тип поиска - рекомендуется использовать X509FindType.FindBySerialNumber.</param>
        /// <param name="findValue">Значение, по которому происходит поиск.</param>
        public void GetMyESignCert(X509FindType findType, object findValue)
        {
            X509Certificate2Collection foundCerts = MyCertsCollection.Find(findType, findValue, true);
            myCert = foundCerts[0];
        }

        /// <summary>
        /// Создаёт электронную подпись и добавляет её в документ
        /// </summary>
        /// <param name="docPath">Полный путь к PDF файлу для подписи</param>
        /// <param name="newDocPath">Полный путь сохранения подписанного файла</param>
        /// <param name="x">Координата X расположения подписи в документе</param>
        /// <param name="y">Координата Н расположения подписи в документе</param>
        /// <returns></returns>
        public String CreateAndPlaceSignature(string docPath, string newDocPath, float x=60, float y=60)
        {
            PdfReader reader = new PdfReader(docPath);
            PdfStamper st = PdfStamper.CreateSignature(reader, new FileStream(newDocPath, FileMode.Create, FileAccess.Write), '\0');
            PdfSignatureAppearance sap = st.SignatureAppearance;

            // Загружаем сертификат в объект iTextSharp
            X509CertificateParser parser = new X509CertificateParser();
            Org.BouncyCastle.X509.X509Certificate[] chain = new Org.BouncyCastle.X509.X509Certificate[] {
                parser.ReadCertificate(this.myCert.RawData)
            };

            sap.Certificate = parser.ReadCertificate(this.myCert.RawData);
            sap.Reason = reason;
            sap.Location = location;

            //sap.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.NAME_AND_DESCRIPTION;
            sap.SignDate = DateTime.Now;

            #region SIGNATURE_VISUAL
            sap.Acro6Layers = true;
            sap.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.DESCRIPTION;

            int pageNum;
            if (signaturePageNum == null || signaturePageNum > reader.NumberOfPages)
                pageNum = reader.NumberOfPages;
            else
                pageNum = (int)signaturePageNum;

            sap.SetVisibleSignature(new iTextSharp.text.Rectangle(x, y, 254, 154), pageNum, "Signature");
            PdfTemplate layer2 = sap.GetLayer(2);
            String text = $"Сертификат: {this.myCert.SerialNumber}\n"
                    + $"Владелец: {Regex.Match(this.myCert.Subject, @"CN=(?<name>[А-Яа-я\s]+),").Groups["name"].Value}\n"
                    + $"Срок действия с {this.myCert.NotBefore.ToString("d")} по {this.myCert.NotAfter.ToString("d")}\n";

            BaseFont baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font font = new Font(baseFont, 5.5F);

            float leading = 10F;
            float MARGIN = 2;

            Rectangle dataRect = new Rectangle(
                    MARGIN,
                    MARGIN,
                    sap.Rect.Width - MARGIN - 2,
                    sap.Rect.Height - MARGIN - 2);

            //Вставляем рамку электронной подписи
            layer2.RoundRectangle(dataRect.Left, dataRect.Bottom, dataRect.Right, dataRect.Top, 10);
            layer2.Stroke();

            //Вставляем лого
            Image image = Image.GetInstance(LogoPath);
            image.SetAbsolutePosition(10F, 60F);
            image.ScaleAbsolute(45F, 30F);
            layer2.AddImage(image);

            //Заголовок Основной 
            String title = "Документ подписан усиленной\nэлектронной цифровой подписью";
            ColumnText ct3 = new ColumnText(layer2);
            ct3.RunDirection = (sap.RunDirection);
            ct3.SetSimpleColumn(new Phrase(title, new Font(baseFont, 7F)), 65F, 87F, 175F, 35F, leading, Element.ALIGN_CENTER);
            ct3.Go();

            //Вставляем полоску сведения
            layer2.Rectangle(10F, 43F, 175F, 15F);
            layer2.SetRGBColorStrokeF(lineColors[0], lineColors[1], lineColors[2]);
            layer2.SetRGBColorFillF(lineColors[0], lineColors[1], lineColors[2]);
            layer2.FillStroke();

            //Заголовок Сведения 
            String info = "Сведения о сертификате ЭП";
            ColumnText ct2 = new ColumnText(layer2);
            ct2.RunDirection = (sap.RunDirection);
            BaseColor color = new BaseColor(255, 255, 255);
            ct2.SetSimpleColumn(new Phrase(info, new Font(baseFont, 10F, 1, color)), 15F, 57F, 175F, 15F, leading, Element.ALIGN_CENTER);
            ct2.Go();

            //Сведения;
            ColumnText ct = new ColumnText(layer2);
            ct.RunDirection = (sap.RunDirection);
            ct.SetSimpleColumn(new Phrase(text, font), 10, 10, 180, 40, leading, Element.ALIGN_LEFT);
            ct.Go();
            #endregion

            // Выбираем подходящий тип фильтра
            PdfName filterName = new PdfName("CryptoPro PDF");

            // Создаем подпись
            PdfSignature dic = new PdfSignature(filterName, PdfName.ADBE_PKCS7_DETACHED);
            dic.Date = new PdfDate(sap.SignDate);
            dic.Name = "PdfPKCS7 signature";
            if (sap.Reason != null)
                dic.Reason = sap.Reason;
            if (sap.Location != null)
                dic.Location = sap.Location;
            sap.CryptoDictionary = dic;

            Dictionary<PdfName, int> hashtable = new Dictionary<PdfName, int>();
            hashtable[PdfName.CONTENTS] = this.buff * 2 + 2;
            sap.PreClose(hashtable);
            Stream s = sap.GetRangeStream();
            MemoryStream ss = new MemoryStream();
            int read = 0;
            byte[] buff = new byte[8192];
            while ((read = s.Read(buff, 0, 8192)) > 0)
            {
                ss.Write(buff, 0, read);
            }

            // Вычисляем подпись
            #region END_OF_SIGN
            byte[] cntnt = ss.ToArray();
            ICertificate2 myCert = GetSignerCertificate(this.myCert.SerialNumber);
            byte[] pk = SignDoc(myCert, cntnt);
            #endregion

            // Помещаем подпись в документ
            byte[] outc = new byte[this.buff];
            PdfDictionary dic2 = new PdfDictionary();
            Array.Copy(pk, 0, outc, 0, pk.Length);
            dic2.Put(PdfName.CONTENTS, new PdfString(outc).SetHexWriting(true));
            sap.Close(dic2);

            return $"Документ {docPath} успешно подписан на ключе {this.myCert.Subject} => {2}.";
        }

        // Метод необходим для создания подписи CAdES XLong1 - штамп времени и доказательства подлинности
        private ICertificate2 GetSignerCertificate(string serial)
        {
            ICertificate2 ecpCert = null;
            string storeName = "My";
            CAPICOM_STORE_LOCATION storeLocation = CAPICOM_STORE_LOCATION.CAPICOM_CURRENT_USER_STORE;
            CAPICOM_STORE_OPEN_MODE openMode = CAPICOM_STORE_OPEN_MODE.CAPICOM_STORE_OPEN_READ_ONLY;

            Store cStore = new Store();
            cStore.Open(storeLocation, storeName, openMode);
            foreach (ICertificate2 cert in cStore.Certificates)
            {
                if (cert.SerialNumber == serial)
                {
                    ecpCert = cert;
                    break;
                }
            }
            cStore.Close();
            return ecpCert;
        }

        // Метод подписывает документ УЭЦП CAdES XLong1 - штамп времени и доказательства подлинности
        private byte[] SignDoc(ICertificate2 cert, byte[] content)
        {
            CPSigner dSigner = new CPSigner();

            dSigner.Certificate = cert;
            dSigner.TSAAddress = TSAAddress;
            dSigner.Options = CAPICOM_CERTIFICATE_INCLUDE_OPTION.CAPICOM_CERTIFICATE_INCLUDE_WHOLE_CHAIN;

            CadesSignedData signedData = new CadesSignedData();
            signedData.Content = content;
            byte[] bSignedData = signedData.SignCades(dSigner, CADESCOM_CADES_TYPE.CADESCOM_CADES_BES, true, CAPICOM_ENCODING_TYPE.CAPICOM_ENCODE_BINARY);
            bSignedData = signedData.EnhanceCades(CADESCOM_CADES_TYPE.CADESCOM_CADES_DEFAULT, TSAAddress, CAPICOM_ENCODING_TYPE.CAPICOM_ENCODE_BINARY);

            return bSignedData;
        }
    }

    /// <summary>
    /// Класс для верификации подписи документа
    /// </summary>
    public class DigitalVerifier
    {
        public Dictionary<string, string> Verify(string docPath)
        {
            Dictionary<string, string> result = new Dictionary<string, string>{ };

            string document = docPath;

            // Открываем документ
            PdfReader reader = new PdfReader(document);

            // Получаем подписи из документа
            AcroFields af = reader.AcroFields;
            List<string> names = af.GetSignatureNames();
            foreach (string name in names)
            {
                result["Name"] = name;
                result["CoversWholeDoc"] = af.SignatureCoversWholeDocument(name).ToString();
                result["Revision"] = af.GetRevision(name) + " of " + af.TotalRevisions;

                // Проверяем подпись
                PdfDictionary singleSignature = af.GetSignatureDictionary(name);
                PdfString asString1 = singleSignature.GetAsString(PdfName.CONTENTS);
                byte[] signatureBytes = asString1.GetOriginalBytes();

                RandomAccessFileOrArray safeFile = reader.SafeFile;

                PdfArray asArray = singleSignature.GetAsArray(PdfName.BYTERANGE);
                using (
                    Stream stream =
                        (Stream)
                        new RASInputStream(
                            new RandomAccessSourceFactory().CreateRanged(
                                safeFile.CreateSourceView(),
                                (IList<long>)asArray.AsLongArray())))
                {
                    using (MemoryStream ms = new MemoryStream((int)stream.Length))
                    {
                        stream.CopyTo(ms);
                        byte[] data = ms.GetBuffer();

                        // Создаем объект ContentInfo по сообщению.
                        // Это необходимо для создания объекта SignedCms.
                        ContentInfo contentInfo = new ContentInfo(data);

                        // Создаем SignedCms для декодирования и проверки.
                        SignedCms signedCms = new SignedCms(contentInfo, true);

                        // Декодируем подпись
                        signedCms.Decode(signatureBytes);

                        bool checkResult;

                        try
                        {
                            signedCms.CheckSignature(true);
                            checkResult = true;
                        }
                        catch (Exception)
                        {
                            checkResult = false;
                        }

                        result["Changed"] = (!checkResult).ToString();

                        foreach (var sinerInfo in signedCms.SignerInfos)
                        {
                            result["Certificate"] = sinerInfo.Certificate.ToString();
                            X509Certificate2 cert = signedCms.Certificates[0];
                            var isCapiValid = cert.Verify();
                            result["CAPI Validation"] = isCapiValid.ToString();


                            foreach (var attribute in sinerInfo.SignedAttributes)
                            {
                                if (attribute.Oid.Value == "1.2.840.113549.1.9.5")
                                {
                                    Pkcs9SigningTime signingTime = attribute.Values[0] as Pkcs9SigningTime;
                                    result["Date"] = signingTime.SigningTime.ToString("G");
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }
    }
}

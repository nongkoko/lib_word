using System.Globalization;

namespace jomiunsWords
{

    public class Terbilang2008
    {
        private string _decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        private string[] _angka = { "", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas" };
        private string[] _level = { "", "ribu", "juta", "miliar" };

        public string Result(double iAngka)
        {
            string aRetval = "";
            string aAngka;
            string aKoma = "";
            string[] aSplit;
            int aTotalLength;

            aSplit = iAngka.ToString(CultureInfo.CurrentCulture).Split(new string[] { _decimalSeparator }, StringSplitOptions.None);
            aAngka = aSplit[0];
            if (aSplit.Length > 1)
            {
                aKoma = aSplit[1];
            }

            aTotalLength = aAngka.Length;
            int aSpent;
            int aLevel = 0;
            string aHasil;
            aRetval = "";
            while (aTotalLength > 0)
            {
                aSpent = Math.Min(aTotalLength, 3);
                aHasil = Spell3Char(aAngka.Substring(Math.Max(aTotalLength - 3, 0), aSpent));
                if (aHasil != "")
                {
                    if (_level[aLevel] != "")
                    {
                        aHasil = aHasil + " " + _level[aLevel];
                    }
                    aHasil = aHasil.Replace("satu ribu", "seribu");
                    if (aRetval == "")
                    {
                        aRetval = aHasil;
                    }
                    else
                    {
                        aRetval = aHasil + " " + aRetval;
                    }
                }

                aTotalLength = aTotalLength - aSpent;
                aLevel += 1;
            }
            if (aRetval == "")
            {
                aRetval = "nol";
            }
            if (aKoma != "")
            {
                aRetval = aRetval + " koma " + SpellKoma(aKoma);
            }
            return aRetval.Trim();
        }

        private string Spell3Char(string iAngka)
        {
            iAngka = iAngka.PadLeft(3, '0');
            string aCharPertama = iAngka.Substring(0, 1);
            string aCharKedua = iAngka.Substring(1, 1);
            string aCharKetiga = iAngka.Substring(2, 1);
            string aRetval = "";

            if (aCharPertama == "1")
            {
                aRetval = "seratus ";
            }
            else
            {
                if (aCharPertama != "0")
                {
                    aRetval = _angka[int.Parse(aCharPertama)] + " ratus ";
                }
            }
            if (aCharKedua == "1")
            {
                aRetval = aRetval + _angka[int.Parse(iAngka.Substring(1, 2))];
            }
            else
            {
                if (aCharKedua != "0")
                {
                    aRetval = aRetval + _angka[int.Parse(aCharKedua)] + " puluh ";
                }
                if (aCharKetiga != "0")
                {
                    aRetval = aRetval + _angka[int.Parse(aCharKetiga)];
                }
            }
            return aRetval.Trim();
        }

        private string SpellKoma(string iAngka)
        {
            string aRetval = "";
            for (int aInt = 0; aInt < iAngka.Length; aInt++)
            {
                if (aRetval == "")
                {
                    aRetval = AngkaToUcapKoma(iAngka.Substring(aInt, 1));
                }
                else
                {
                    aRetval = aRetval + " " + AngkaToUcapKoma(iAngka.Substring(aInt, 1));
                }
            }
            return aRetval;
        }

        private string AngkaToUcapKoma(string iString)
        {
            if (iString == "0")
            {
                return "kosong";
            }
            else
            {
                return _angka[int.Parse(iString)];
            }
        }
    }
}

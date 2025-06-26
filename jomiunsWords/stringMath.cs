namespace jomiunsWords
{
    public class clsMathInString
    {
        private const int _akurasi = 15;
        private const string _siKoma = ".";
        private int[,] _arrayHasil = new int[1, 1];
        private int _varSisa;

        private int zGetSisa()
        {
            int aTemp = _varSisa;
            _varSisa = 0;
            return aTemp;
        }

        private void zPutSisa(int berapa)
        {
            _varSisa = berapa;
        }

        private static string zTanda(string angkaString1, string angkaString2)
        {
            string retVal = "-";
            if ((angkaString1.Contains("-") && angkaString2.Contains("-")) ||
                (!angkaString1.Contains("-") && !angkaString2.Contains("-")))
            {
                retVal = "";
            }
            return retVal;
        }

        /// <summary>
        /// Supports unlimited integer multiplication, e.g. 17291851986982691586668256192568215 * 5687659286592685928561925681925689
        /// </summary>
        public string perkalianStringInteger(string string1, string string2)
        {
            int len1, len2, kolomCell, barisCell, aPuluhan;
            string aTempString = "";
            string tandaAkhir = zTanda(string1, string2);

            string1 = string1.Replace(_siKoma, "").Replace("-", "");
            string2 = string2.Replace(_siKoma, "").Replace("-", "");

            len1 = string1.Length;
            len2 = string2.Length;

            _arrayHasil = new int[len2 + 1, len1 + len2 + 2];

            for (int indexBawah = len2; indexBawah >= 1; indexBawah--)
            {
                int angkaBawah = int.Parse(string2.Substring(indexBawah - 1, 1));
                kolomCell = len1 + indexBawah;
                for (int indexAtas = len1; indexAtas >= 1; indexAtas--)
                {
                    int angkaAtas = int.Parse(string1.Substring(indexAtas - 1, 1));
                    int hasilPerkalian = zGetSisa() + angkaAtas * angkaBawah;
                    int satuan = hasilPerkalian % 10;
                    aPuluhan = hasilPerkalian / 10;
                    zPutSisa(aPuluhan);
                    _arrayHasil[indexBawah, kolomCell] = satuan;
                    kolomCell = kolomCell - 1;
                }
                if (_varSisa != 0)
                {
                    _arrayHasil[indexBawah, kolomCell] = zGetSisa();
                }
            }

            for (kolomCell = len1 + len2; kolomCell >= 1; kolomCell--)
            {
                int hasilPenambahan = zGetSisa();
                for (barisCell = len2; barisCell >= 1; barisCell--)
                {
                    hasilPenambahan = hasilPenambahan + _arrayHasil[barisCell, kolomCell];
                }
                int satuan = hasilPenambahan % 10;
                aPuluhan = hasilPenambahan / 10;
                zPutSisa(aPuluhan);
                aTempString = satuan.ToString() + aTempString;
            }

            // Remove leading zeros
            int inta = 0;
            while (inta < aTempString.Length && aTempString[inta] == '0')
            {
                inta++;
            }
            if (inta < aTempString.Length)
                aTempString = aTempString.Substring(inta);
            else
                aTempString = "0";

            return tandaAkhir + aTempString;
        }

        public static string penambahanStringInteger(string iInteger1, string iInteger2)
        {
            int aJumlahMax = Math.Max(iInteger1.Length, iInteger2.Length);

            iInteger1 = iInteger1.PadLeft(aJumlahMax, '0');
            iInteger2 = iInteger2.PadLeft(aJumlahMax, '0');

            int aSisa = 0;
            string aHasil = "";
            for (int aInt = aJumlahMax - 1; aInt >= 0; aInt--)
            {
                int aAtas = int.Parse(iInteger1.Substring(aInt, 1));
                int aBawah = int.Parse(iInteger2.Substring(aInt, 1));
                int aTempor = (aAtas + aBawah) + aSisa;
                aHasil = (aTempor % 10).ToString() + aHasil;
                aSisa = aTempor / 10;
            }
            if (aSisa == 0)
            {
                return aHasil;
            }
            else
            {
                return aSisa.ToString() + aHasil;
            }
        }

        public string xPangkatYinteger(string iEx, int iYe)
        {
            string aHasil = "1";
            for (int aInt = 1; aInt <= iYe; aInt++)
            {
                aHasil = perkalianStringInteger(aHasil, iEx);
            }
            return aHasil;
        }
    }
}

using System;

namespace ATP_Common_Plugin.Utils
{
    class dictionaryGUID
    {
        // ФОП
        public const string SharedParameterFilePath = @"P:\MOS-TLP\GROUPS\ALLGEMEIN\02_ATP_STANDARDS\07_BIM\01_Settings\07_Shared parameter lists\ATP_ФОП.txt";

        // Парамтеры ADSK
        public static Guid ADSKGroup = new Guid("3de5f1a4-d560-4fa8-a74f-25d250fb3401");
        public static Guid ADSKKomp = new Guid("8dd021be-382d-4776-afd4-75996e351de3"); // ADSK_Комплект
        public static Guid ADSKName = new Guid("e6e0f5cd-3e26-485b-9342-23882b20eb43");
        public static Guid ADSKMark = new Guid("2204049c-d557-4dfc-8d70-13f19715e46d");
        public static Guid ADSKSign = new Guid("9c98831b-9450-412d-b072-7d69b39f4029"); // ADSK_Обозначение
        public static Guid ADSKManufactory = new Guid("a8cdbf7b-d60a-485e-a520-447d2055f351");
        public static Guid ADSKFabricName = new Guid("87ce1509-068e-400f-afab-75df889463c7");
        public static Guid ADSKUnit = new Guid("4289cb19-9517-45de-9c02-5a74ebf5c86d");
        public static Guid ADSKCount = new Guid("8d057bb3-6ccd-4655-9165-55526691fe3a");
        public static Guid ADSKThicknes = new Guid("381b467b-3518-42bb-b183-35169c9bdfb3");
        public static Guid ADSKSizeArea = new Guid("b6a46386-70e9-4b1f-9fdb-8e1e3f18a673");
        public static Guid ADSKSlope = new Guid("80d34cef-47d0-4c21-8409-6305fda3a286");

        // Парамтеры ATP
        public static Guid ATPMarkScriot = new Guid("bf962560-877e-4870-b390-a649ac64bf7c");
        public static Guid ATPHost = new Guid("6dfb04f9-1cdb-4f00-988b-25ce488d52f0");
    }
}

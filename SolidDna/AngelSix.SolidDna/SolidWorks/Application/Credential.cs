using System;

namespace AngelSix.SolidDna
{
    internal class Credential
    {
        private const string API_LICENCE_KEY_FILE = "solidworks_api_license_key.txt";

        /// <summary>
        ///  Get raw solidworks api license key from file solidworks_api_license_key.txt
        ///  File is located in generated bin/x64/Debug
        ///  TODO: It would be better to store the hashed key
        /// </summary>
        /// <returns></returns>
        static public string getSolidWorksLicenseAPIKey()
        {
            var key = System.IO.File.ReadAllText(API_LICENCE_KEY_FILE);
            return key;
        }
    }
}

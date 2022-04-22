using System;
using System.IO;
using System.Reflection;

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
        static public string GetSolidWorksLicenseAPIKey()
        {
            // TODO: Do not store key on client pc
            // instead send the api key via https if user has logged in successfully
            // then only hold the api key in RAM as long as the app is running
            // use ProtectedMemory to securly store the key in RAM
            var key = File.ReadAllText(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + API_LICENCE_KEY_FILE);
            return key;
        }
    }
}

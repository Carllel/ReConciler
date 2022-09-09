using System;
using System.IO;

namespace ReConciler
{
    public class JSONReadWrite
    {
        public JSONReadWrite()
        { }

        public string Read(string fileName)
        {

            string jsonResult;

            try
            {
                using (StreamReader streamReader = new StreamReader(fileName))
                {
                    jsonResult = streamReader.ReadToEnd();
                }
                return jsonResult;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public void Write(string fileName, string jSONString)
        {
            try
            {
                using (var streamWriter = File.CreateText(fileName))
                {
                    streamWriter.Write(jSONString);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}

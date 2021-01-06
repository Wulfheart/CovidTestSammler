using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace CovidTestOutlookCollector
{
    class Config
    {
        public CovidIgnore Data { get; set; }

        private string path { get; set; }

        public Config(string folder)
        {
            path = Path.Combine(folder, "covidignore.json");
            FileStream fs = File.Open(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            
            
            StreamReader sr = new StreamReader(fs, Encoding.UTF8);

            string covidJson = sr.ReadToEnd();
            sr.Close();
            fs.Close();
            File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);
            try
            {
                Data = JsonSerializer.Deserialize<CovidIgnore>(covidJson);

            }
            catch (System.Text.Json.JsonException ex)
            {
                Data = new CovidIgnore();
                Save();
            }
        }

        public void AddFile(string file)
        {
            Data.files.Add(file);
        }

        public void AddID(string ID)
        {
            Data.ids.Add(ID);
        }

        public void Save()
        {
            FileStream fs = File.Open(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);


            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            sw.Write(JsonSerializer.Serialize(Data));

            
            sw.Close();
            fs.Close();
            
        }

        public void Reload()
        {
            throw new NotImplementedException();
        }
    }
}

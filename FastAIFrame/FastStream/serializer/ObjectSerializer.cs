using MessagePack;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace FastStream.serializer
{
   public class ObjectSerializer
    {
        JsonSerializerSettings settings = new JsonSerializerSettings();
       
        public byte[] Serializer<T> (T obj)
        {
            return MessagePackSerializer.Serialize<T>(obj);
          
            // Pack obj to stream.
            //serializer.Pack(stream, obj);
        }
        public T Deserialize<T>(byte[] data)
        {
          return  MessagePackSerializer.Deserialize<T>(data);
        }

        public string SerializerJSON<T>(T obj)
        {
            settings.DateFormatString = "yyyy-MM-dd HH:mm:ss";
            return    JsonConvert.SerializeObject(obj, settings);
        }

        public T DeserializeJSON<T>(string json)
        {
           return JsonConvert.DeserializeObject<T>(json);
        }

       

    }
}

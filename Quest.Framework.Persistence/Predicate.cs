using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Quest.Framework.Persistance
{
    public class QueryParameters : Dictionary<String, Object>
    {
        public QueryParameters(Dictionary<String, Object> keyValues)
        {
            //foreach (String key in keyValues.Keys)
            //{
            //    if (string.IsNullOrEmpty(result))
            //        result = string.Format(" WHERE {0} = {1}", key, FormatValue(this[key]));
            //    else
            //        result += string.Format(" AND {0} = {1}", key, FormatValue(this[key]));
            //}
            //this.Add
        }
        public override string ToString()
        {
            String result = "";
            foreach (String key in this.Keys)
            {
                if (string.IsNullOrEmpty(result))
                    result = string.Format(" WHERE {0} = {1}", key, FormatValue(this[key]));
                else
                    result += string.Format(" AND {0} = {1}", key, FormatValue(this[key]));
            }

            return result;
        }
        private string FormatValue(Object value)
        {
            if(value is string)
            {
                return string.Format("\"{0}\"", value);
            }
            if (value is DateTime)
            {
                return ((DateTime) value).ToString("yyyy-MM-dd");
            }
            if(value is int)
            {
                return ((int)value).ToString();
            }
            if(value is decimal)
            {
                return ((decimal)value).ToString();
            }
            return "";
        }
    }
}

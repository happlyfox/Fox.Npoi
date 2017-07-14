using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace Fox.Npoi.Ext
{
    public static class ObjectExtensions
    {
        public static PropertyInfo[] GetProperties<T>(this T entity)
        {
            return entity == null ? new PropertyInfo[0] : entity.GetType().GetProperties();
        }

        public static string GetDescription(this PropertyInfo propertyInfo)
        {
            var attribute = propertyInfo.GetCustomAttributes(typeof(DescriptionAttribute), true).FirstOrDefault();
            return attribute == null ? propertyInfo.Name : ((DescriptionAttribute)attribute).Description;
        }

        public static dynamic GetValue<T>(this T entity, PropertyInfo propertyInfo)
        {
            return propertyInfo.GetValue(entity, null);
        }
    }
}
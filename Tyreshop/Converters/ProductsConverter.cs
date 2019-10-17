using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using Tyreshop.Utils;

namespace Tyreshop.Converters
{
    class ProductsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var values = value as IEnumerable;
            var type = parameter as Type;

            if (values == null || type == null)
                return null;

            if (type.GetInterfaces().Any(x => x == typeof(IParamsWrapper)))
            {
                var instance = (IParamsWrapper)type.Assembly.CreateInstance(type.FullName);
                instance.Pars = (IEnumerable<BdProducts>)value;
                //returned value should be IEnumerable with one element. 
                //Otherwise we will not see children nodes
                return new List<IParamsWrapper> { instance };
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

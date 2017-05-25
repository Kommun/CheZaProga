using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CheZaProgaV2.Controls
{
    public class WithoutMatch : Control
    {
        public string SourceAddress
        {
            get { return (string)this.GetValue(SourceAddressProperty); }
            set { this.SetValue(SourceAddressProperty, value); }
        }
        public static readonly DependencyProperty SourceAddressProperty = DependencyProperty.Register(
          "SourceAddress", typeof(string), typeof(WithoutMatch), null);

        static WithoutMatch()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(WithoutMatch), new FrameworkPropertyMetadata(typeof(WithoutMatch)));
        }
    }
}

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
using CheZaProgaV2;

namespace CheZaProgaV2.Controls
{
    public class MultiMatch : Control
    {
        public SourceAddress SourceAddress
        {
            get { return (SourceAddress)this.GetValue(SourceAddressProperty); }
            set { this.SetValue(SourceAddressProperty, value); }
        }
        public static readonly DependencyProperty SourceAddressProperty = DependencyProperty.Register(
          "SourceAddress", typeof(SourceAddress), typeof(MultiMatch), null);

        public List<ResultAddress> ResultAddresses
        {
            get { return (List<ResultAddress>)this.GetValue(ResultAddressesProperty); }
            set { this.SetValue(ResultAddressesProperty, value); }
        }
        public static readonly DependencyProperty ResultAddressesProperty = DependencyProperty.Register(
          "ResultAddresses", typeof(List<ResultAddress>), typeof(MultiMatch), null);

        static MultiMatch()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(MultiMatch), new FrameworkPropertyMetadata(typeof(MultiMatch)));
        }
    }
}

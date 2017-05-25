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
    public class SingleMatch : Control
    {
        public bool IsChecked
        {
            get { return (bool)this.GetValue(IsCheckedProperty); }
            set { this.SetValue(IsCheckedProperty, value); }
        }
        public static readonly DependencyProperty IsCheckedProperty = DependencyProperty.Register(
          "IsChecked", typeof(bool), typeof(SingleMatch), new PropertyMetadata(true));

        public SourceAddress SourceAddress
        {
            get { return (SourceAddress)this.GetValue(SourceAddressProperty); }
            set { this.SetValue(SourceAddressProperty, value); }
        }
        public static readonly DependencyProperty SourceAddressProperty = DependencyProperty.Register(
          "SourceAddress", typeof(SourceAddress), typeof(SingleMatch), null);


        public ResultAddress ResultAddress
        {
            get { return (ResultAddress)this.GetValue(ResultAddressProperty); }
            set { this.SetValue(ResultAddressProperty, value); }
        }
        public static readonly DependencyProperty ResultAddressProperty = DependencyProperty.Register(
          "ResultAddress", typeof(ResultAddress), typeof(SingleMatch), null);

        static SingleMatch()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(SingleMatch), new FrameworkPropertyMetadata(typeof(SingleMatch)));
        }
    }
}

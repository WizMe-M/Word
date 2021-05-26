using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using System.IO;
using Path = System.IO.Path;

namespace Word
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            FontFamily.ItemsSource = Fonts.SystemFontFamilies.OrderBy(f => f.Source);
            FontSize.ItemsSource = new List<double>() { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void SaveExit_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog SavExt = new SaveFileDialog();
            SavExt.Filter = " (*.txt)|*.txt|All files (*.*)|*.*";
            if (SavExt.ShowDialog() == true)
            {
                TextRange doc = new TextRange(A4.Document.ContentStart, A4.Document.ContentEnd);
                using (FileStream fs = File.Create(SavExt.FileName))
                {
                    if (Path.GetExtension(SavExt.FileName).ToLower() == ".txt")
                        doc.Save(fs, DataFormats.Text);
                }
            }
            Application.Current.Shutdown();
        }

        private void Open_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog Open = new OpenFileDialog();
            Open.Filter = "(*.txt)|*.txt|All files (*.*)|*.*";
            if (Open.ShowDialog() == true)
            {
                FileStream fileStream = new FileStream(Open.FileName, FileMode.Open);
                TextRange range = new TextRange(A4.Document.ContentStart, A4.Document.ContentEnd);
                range.Load(fileStream, DataFormats.Text);
            }
        }

        private void Save_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            SaveFileDialog Save = new SaveFileDialog();
            Save.Filter = "(*.txt)|*.txt|All files (*.*)|*.*";
            if (Save.ShowDialog() == true)
            {
                FileStream fileStream = new FileStream(Save.FileName, FileMode.Create);
                TextRange range = new TextRange(A4.Document.ContentStart, A4.Document.ContentEnd);
                range.Save(fileStream, DataFormats.Text);
            }
        }

        private void Print_Executed(object sender, RoutedEventArgs e)
        {
            PrintDialog pd = new PrintDialog();
            if ((pd.ShowDialog() == true))
            {
                FlowDocument doc = new FlowDocument();

                double pageHeight = doc.PageHeight;
                double pageWidth = doc.PageWidth;
                Thickness pagePadding = doc.PagePadding;
                double columnGap = doc.ColumnGap;
                double columnWidth = doc.ColumnWidth;

                doc.PageHeight = pd.PrintableAreaHeight;
                doc.PageWidth = pd.PrintableAreaWidth;
                doc.PagePadding = new Thickness(0);

                doc.ColumnGap = 25;
                doc.ColumnWidth = (doc.PageWidth - doc.ColumnGap
                    - doc.PagePadding.Left - doc.PagePadding.Right) / 2;

                doc.PageHeight = pageHeight;
                doc.PageWidth = pageWidth;
                doc.PagePadding = pagePadding;
                doc.ColumnGap = columnGap;
                doc.ColumnWidth = columnWidth;
                //pd.PrintVisual(A4 as Visual, "Print Visual");
                pd.PrintDocument((((IDocumentPaginatorSource)A4.Document).DocumentPaginator), "printing as paginator");
            }

        }

        private void FontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FontFamily.SelectedItem != null)
                A4.Selection.ApplyPropertyValue(Inline.FontFamilyProperty, FontFamily.SelectedItem);
        }

        private void FontSize_TextChanged(object sender, TextChangedEventArgs e)
        {
            A4.Selection.ApplyPropertyValue(Inline.FontSizeProperty, FontSize.Text);
        }

        private void A4_SelectionChanged(object sender, RoutedEventArgs e)
        {
            object temp = A4.Selection.GetPropertyValue(Inline.FontWeightProperty);
            Bold.IsChecked = (temp != DependencyProperty.UnsetValue) && (temp.Equals(FontWeights.Bold));
            temp = A4.Selection.GetPropertyValue(Inline.FontStyleProperty);
            Italic.IsChecked = (temp != DependencyProperty.UnsetValue) && (temp.Equals(FontStyles.Italic));
            temp = A4.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            Underline.IsChecked = (temp != DependencyProperty.UnsetValue) && (temp.Equals(TextDecorations.Underline));

            temp = A4.Selection.GetPropertyValue(Inline.FontFamilyProperty);
            FontFamily.SelectedItem = temp;
            temp = A4.Selection.GetPropertyValue(Inline.FontSizeProperty);
            FontSize.Text = temp.ToString();
        }

        private void ColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            Color color = FontColorPicker.SelectedColor ?? Colors.Black;
            RTBApplyProperty(A4, TextElement.ForegroundProperty, new SolidColorBrush(color));
        }

        void RTBApplyProperty(RichTextBox richTextBox, DependencyProperty property, object propertyValue)
        {
            var selection = A4?.Selection;
            if (selection != null && propertyValue != null)
                selection.ApplyPropertyValue(property, propertyValue);
        }
    }
}
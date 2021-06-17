using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace Word
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            FontFamily.ItemsSource = Fonts.SystemFontFamilies.OrderBy(f => f.Source);
            FontFamily.SelectedIndex = 21;
            FontSize.ItemsSource = new List<double>() { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            FontSize.SelectedIndex = 4;
        }

        private void SaveCommandBinding(object sender, ExecutedRoutedEventArgs e)
        {
            SaveFileDialog SaveDialog = new SaveFileDialog
            {
                Filter = "(*.txt)|*.txt|All files (*.*)|*.*"
            };
            if ((bool)SaveDialog.ShowDialog())
            {
                FileStream fileStream = new FileStream(SaveDialog.FileName, FileMode.Create);
                TextRange range = new TextRange(Text.Document.ContentStart, Text.Document.ContentEnd);
                range.Save(fileStream, DataFormats.Text);
            }
        }
        private void OpenCommandBinding(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog OpenDialog = new OpenFileDialog
            {
                Filter = "(*.txt)|*.txt|All files (*.*)|*.*"
            };
            if ((bool)OpenDialog.ShowDialog())
            {
                FileStream fileStream = new FileStream(OpenDialog.FileName, FileMode.Open);
                TextRange range = new TextRange(Text.Document.ContentStart, Text.Document.ContentEnd);
                range.Load(fileStream, DataFormats.Text);
            }
            
        }
        private void PrintCommandBinding(object sender, ExecutedRoutedEventArgs e)
        {
            PrintDialog PrintDialog = new PrintDialog();
            if ((bool)PrintDialog.ShowDialog())
                PrintDialog.PrintDocument((((IDocumentPaginatorSource)Text.Document).DocumentPaginator), "printing as paginator");           
        }

        private void FontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(FontFamily.SelectedItem != null)
                Text.Selection.ApplyPropertyValue(Inline.FontFamilyProperty, FontFamily.SelectedItem);
        }

        private void FontSize_TextChanged(object sender, TextChangedEventArgs e)
        {
            Text.Selection.ApplyPropertyValue(Inline.FontSizeProperty, FontSize.Text);
        }

        private void SaveExit_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog SaveDialog = new SaveFileDialog
            {
                Filter = " (*.txt)|*.txt|All files (*.*)|*.*"
            };
            if ((bool)SaveDialog.ShowDialog())
            {
                TextRange file = new TextRange(Text.Document.ContentStart, Text.Document.ContentEnd);
                using (FileStream fileStream = File.Create(SaveDialog.FileName))
                {
                    if (Path.GetExtension(SaveDialog.FileName).ToLower() == ".txt")
                        file.Save(fileStream, DataFormats.Text);
                }
            }
            Application.Current.Shutdown();
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void Text_SelectionChanged(object sender, RoutedEventArgs e)
        {
            //объект, который будет хранить различные настройки выделенного текста
            object textSelection;

            //жирный
            textSelection = Text.Selection.GetPropertyValue(Inline.FontWeightProperty);
            Bold.IsChecked = (textSelection != DependencyProperty.UnsetValue) && (textSelection.Equals(FontWeights.Bold));

            //курсив
            textSelection = Text.Selection.GetPropertyValue(Inline.FontStyleProperty);
            Italic.IsChecked = (textSelection != DependencyProperty.UnsetValue) && (textSelection.Equals(FontStyles.Italic));

            //подчеркивание
            textSelection = Text.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            Underline.IsChecked = (textSelection != DependencyProperty.UnsetValue) && (textSelection.Equals(TextDecorations.Underline));

            //шрифт
            textSelection = Text.Selection.GetPropertyValue(Inline.FontFamilyProperty);
            FontFamily.SelectedItem = textSelection;

            //размер шрифта
            textSelection = Text.Selection.GetPropertyValue(Inline.FontSizeProperty);
            FontSize.Text = textSelection.ToString();

            //маркировка и нумерация
            {
                Paragraph start = Text.Selection.Start.Paragraph;
                Paragraph end = Text.Selection.End.Paragraph;
                if (end != null && start.Parent is ListItem)
                {
                    TextMarkerStyle markerStyle = ((ListItem)end.Parent).List.MarkerStyle;
                    if (markerStyle == TextMarkerStyle.Disc)
                    {
                        ToggleB.IsChecked = true;
                    }
                    else if (markerStyle == TextMarkerStyle.Decimal)
                    {
                        ToggleN.IsChecked = true;
                    }
                }
                else
                {
                    ToggleB.IsChecked = false;
                    ToggleN.IsChecked = false;
                }
            }
        }
    }
}

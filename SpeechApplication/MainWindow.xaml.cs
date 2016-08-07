using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
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
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace SpeechApplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        // Single Command # SpeechRunner 1.0 # Required to Reactivate SpeechRunner 2.0
        private SpeechSynthesizer primairSynth = new SpeechSynthesizer();
        private PromptBuilder primairBuilder = new PromptBuilder();
        private SpeechRecognitionEngine primairEngine = new SpeechRecognitionEngine();
        private Choices primairChoices = new Choices();
        private Grammar primairGram = null;


        // MultiCommand # Speechrunner 2.0
        private SpeechSynthesizer sSynth = new SpeechSynthesizer();
        private PromptBuilder pBuilder = new PromptBuilder();
        private SpeechRecognitionEngine sRecognize = new SpeechRecognitionEngine();
        private Choices sList = new Choices();
        private Grammar gr = null;

        public MainWindow()
        {
            primairSynth.SelectVoiceByHints(VoiceGender.Male);
            primairChoices.Add(new string[] { "Start Listening", "Stop Listening" });
            GrammarBuilder primairGB = new GrammarBuilder(primairChoices);
            primairGB.Culture = Thread.CurrentThread.CurrentCulture;
            primairGram = new Grammar(primairGB);

            try
            {
                sRecognize.RequestRecognizerUpdate();
                sRecognize.LoadGrammar(gr);
                sRecognize.SpeechRecognized += PrimairRecognize_SpeechRecognized;
                sRecognize.SetInputToDefaultAudioDevice();
                sRecognize.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message);
            }
            
            InitializeComponent();
            sSynth.SelectVoiceByHints(VoiceGender.Male);
            sList.Add(new string[] { 
                "Stop Chrome",
                "Boot Visual Studio",
                "Start Calculator",
                "Start Chrome",
                "Lock",
                "Youtube Regular",
                "Boot Notepad++",
                "Boot Eclipse",
                "Start Verkenner",
                "Start Discord",
                "Screenshot",
                "Skype",
                "Visual Studio Code",
                "Start Taskmanager",
                "Stop Chrome"
            });
            GrammarBuilder gb = new GrammarBuilder(sList);
            gb.Culture = Thread.CurrentThread.CurrentCulture;
            gr = new Grammar(gb);
            textRichTextBox.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            textRichTextBox.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;

           
        }

        void PrimairRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch(e.Result.Text.ToString())
            {
                case "Start Listening":
                    BootSpeech();
                    break;
                case "Stop Listening":
                    sRecognize.RecognizeAsyncStop();
                    break;
            }
        }

        void recog_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            MessageBox.Show("Speech recognized: " + e.Result.Text);
        }

        
        private void speechToTextButton_Click(object sender, RoutedEventArgs e)
        {
            pBuilder.ClearContent();
            pBuilder.AppendText(new TextRange(textRichTextBox.Document.ContentStart,textRichTextBox.Document.ContentEnd).Text);
            sSynth.Speak(pBuilder);
            pBuilder.ClearContent();
        }

        private void BootSpeech()
        {
            startbutton.IsEnabled = false;
            Stopbutton.IsEnabled = true;

            try
            {
                sRecognize.RequestRecognizerUpdate();
                sRecognize.LoadGrammar(gr);
                sRecognize.SpeechRecognized += PrimairRecognize_SpeechRecognized;
                sRecognize.SetInputToDefaultAudioDevice();
                sRecognize.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message);
            }
        }

        private void startbutton_Click(object sender, RoutedEventArgs e)
        {
            startbutton.IsEnabled = false;
            Stopbutton.IsEnabled = true;
            
            try
            {
                sRecognize.RequestRecognizerUpdate();
                sRecognize.LoadGrammar(gr);
                sRecognize.SpeechRecognized += sRecognize_SpeechRecognized;
                sRecognize.SetInputToDefaultAudioDevice();
                sRecognize.RecognizeAsync(RecognizeMode.Multiple);                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message);
            }
        }

        private void returnAnswer(PromptBuilder builder, SpeechSynthesizer synth, string response)
        {
            builder.ClearContent();
            pBuilder.AppendText(response);
            synth.Speak(builder);
            builder.ClearContent();       
        }


        void sRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text.ToString())
            {
                case "Start Calculator":
                    Process.Start("calc");
                    returnAnswer(pBuilder, sSynth, "Calculator Started");
                    break;
                case "Start Chrome":
                    Process.Start("Chrome");
                    returnAnswer(pBuilder, sSynth, "Chrome Browser started");
                    break;
                case "Lock":
                    LockWorkStation();
                    break;
                case "Youtube Regular":
                    Process.Start("Chrome", "https://www.youtube.com/watch?v=JRfuAukYTKg&index=1&list=PLPhF6hYLIIMi5hj-ZiCyHm52akRDDpRNC");
                    returnAnswer(pBuilder, sSynth, "Regular Playlist started");
                    break;         
                case "Boot Visual Studio":
                    Process.Start("devenv");
                    returnAnswer(pBuilder, sSynth, "Visual Studio Started.");
                    break;
                case "Boot Notepad++":
                    Process.Start("Notepad++");
                    returnAnswer(pBuilder, sSynth, "Notepad++ Started.");
                    break;
                case "Boot Eclipse":
                    Process.Start(@"C:\Users\11400747\eclipse\java-mars\eclipse\eclipse.exe");
                    returnAnswer(pBuilder, sSynth, "Eclipse Started.");
                    break;
                case "Start Verkenner":
                    Process.Start("explorer");
                    returnAnswer(pBuilder, sSynth, "Explorer Opened.");
                    break;
                case "Start Discord":
                    Process.Start(@"C:\Users\11400747\AppData\Local\Discord\Update.exe");
                    returnAnswer(pBuilder, sSynth, "Discord Started.");
                    break;
                case "Screenshot":
                    Process.Start(@"C:\Program Files (x86)\Skillbrains\lightshot\Lightshot.exe");
                    returnAnswer(pBuilder, sSynth, "Lightshot started.");
                    break;
                case "Skype":
                    Process.Start("Skype");
                    returnAnswer(pBuilder, sSynth, "Skype Booted.");
                    break;
                case "Visual Studio Code":
                    Process.Start("Code");
                    returnAnswer(pBuilder, sSynth, "Visual Studio Code Started.");
                    break;
                case "Start Taskmanager":
                    Process.Start("Taskmgr");
                    returnAnswer(pBuilder, sSynth, "Task Manager activated.");
                    break;
                case "Stop Chrome":
                   Process[] proc = Process.GetProcessesByName("Chrome");
                   foreach(Process p in proc)
                    {
                        p.Kill();
                    }
                   returnAnswer(pBuilder, sSynth, "Chrome Processes terminated.");
                   break;
                   
                default:
                    returnAnswer(pBuilder, sSynth, "Command Not found.");
                    break;
            }
        }

        private void Stopbutton_Click(object sender, RoutedEventArgs e)
        {
            sRecognize.RecognizeAsyncStop();
            Stopbutton.IsEnabled = false;
            startbutton.IsEnabled = true;
        }

        [DllImport("user32")]
        public static extern void LockWorkStation();
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeechApplication
{
    public class Spraak2 : Spraak
    {

        private SpeechSynthesizer Synth = new SpeechSynthesizer();
        private PromptBuilder builder = new PromptBuilder();
        private SpeechRecognitionEngine engine = null;
        private Choices choices = new Choices();
        private GrammarBuilder gramBuilder = null;
        private Grammar gram = null;

        public Spraak2()
        {
            Func();
        }

        private void Func()
        {
            Synth.SelectVoiceByHints(VoiceGender.Male);
            choices.Add(new string[]{
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
            });
            engine = new SpeechRecognitionEngine();
            gramBuilder = new GrammarBuilder(choices);
            gramBuilder.Culture = Thread.CurrentThread.CurrentCulture;
            gram = new Grammar(gramBuilder);
            engine.LoadGrammar(gram);
            ReturnAnswer(builder, Synth, "Spraakprogramma 2 started.");
        }
        public void Start()
        {
            try
            {
                Func();

                engine.RequestRecognizerUpdate();
                engine.SpeechRecognized += engine_SpeechRecognized;
                engine.SetInputToDefaultAudioDevice();
                engine.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message);
            }
        }

        public void Stop()
        {
            try
            {
                engine.RecognizeAsyncStop();
                ReturnAnswer(builder, Synth, "Programma 2 stopped.");
                engine = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message);
            }
        }

        [DllImport("user32")]
        public static extern void LockWorkStation();

        void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text.ToString())
            {
                case "Start Calculator":
                    Process.Start("calc");
                    //ReturnAnswer(builder, Synth, "Calculator Started");
                    break;
                case "Start Chrome":
                    Process.Start("Chrome");
                    ReturnAnswer(builder, Synth, "Chrome Browser started");
                    break;
                case "Lock":
                    LockWorkStation();
                    break;
                case "Youtube Regular":
                    Process.Start("Chrome", "https://www.youtube.com/watch?v=JRfuAukYTKg&index=1&list=PLPhF6hYLIIMi5hj-ZiCyHm52akRDDpRNC");
                    ReturnAnswer(builder, Synth, "Regular Playlist started");
                    break;
                case "Boot Visual Studio":
                    Process.Start("devenv");
                    ReturnAnswer(builder, Synth, "Visual Studio Started.");
                    break;
                case "Boot Notepad++":
                    Process.Start("Notepad++");
                    ReturnAnswer(builder, Synth, "Notepad++ Started.");
                    break;
                case "Boot Eclipse":
                    Process.Start(@"C:\Users\11400747\eclipse\java-mars\eclipse\eclipse.exe");
                    ReturnAnswer(builder, Synth, "Eclipse Started.");
                    break;
                case "Start Verkenner":
                    Process.Start("explorer");
                    ReturnAnswer(builder, Synth, "Explorer Opened.");
                    break;
                case "Start Discord":
                    Process.Start(@"C:\Users\11400747\AppData\Local\Discord\Update.exe");
                    ReturnAnswer(builder, Synth, "Discord Started.");
                    break;
                case "Screenshot":
                    Process.Start(@"C:\Program Files (x86)\Skillbrains\lightshot\Lightshot.exe");
                    ReturnAnswer(builder, Synth, "Lightshot started.");
                    break;
                case "Skype":
                    Process.Start("Skype");
                    ReturnAnswer(builder, Synth, "Skype Booted.");
                    break;
                case "Visual Studio Code":
                    Process.Start("Code");
                    ReturnAnswer(builder, Synth, "Visual Studio Code Started.");
                    break;
                case "Start Taskmanager":
                    Process.Start("Taskmgr");
                    ReturnAnswer(builder, Synth, "Task Manager activated.");
                    break;
                case "Stop Chrome":
                    Process[] proc = Process.GetProcessesByName("Chrome");
                    foreach (Process p in proc)
                    {
                        p.Kill();
                    }
                    ReturnAnswer(builder, Synth, "Chrome Processes terminated.");
                    break;

                default:
                    ReturnAnswer(builder, Synth, "Command Not found.");
                    break;
            }
        }
    }
}

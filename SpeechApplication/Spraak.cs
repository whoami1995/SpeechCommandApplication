using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpeechApplication
{
    public class Spraak
    {

        private SpeechSynthesizer Synth = new SpeechSynthesizer();
        private PromptBuilder builder = new PromptBuilder();
        private SpeechRecognitionEngine engine = new SpeechRecognitionEngine();
        private Choices choices = new Choices();
        private GrammarBuilder gramBuilder = null;
        private Grammar gram = null;
        private Spraak2 program = null;
        private bool booted = false;

        public void Func()
        {
            program = new Spraak2();
            Synth.SelectVoiceByHints(VoiceGender.Male);
            choices.Add(new string[] { "Boot", "Shut" });
            gramBuilder = new GrammarBuilder(choices);
            gramBuilder.Culture = Thread.CurrentThread.CurrentCulture;
            gram = new Grammar(gramBuilder);

            try
            {
                engine.RequestRecognizerUpdate();
                engine.LoadGrammar(gram);
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
        void engine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch(e.Result.Text.ToString())
            {
                case "Boot":
                    if(!booted)
                    {
                        ReturnAnswer(builder, Synth, "Activating Recognizer 2");
                        program.Start();
                        booted = true;
                    }      
                    break;
                case "Shut":
                    if(booted)
                    {
                        ReturnAnswer(builder, Synth, "Deactivating Recognizer 2");
                        program.Stop();
                        booted = false;
                    }
                    break;
            }
        }

        public void ReturnAnswer(PromptBuilder builder, SpeechSynthesizer synth, string response)
        {
            try
            {
                builder.ClearContent();
                builder.AppendText(response);
                synth.Speak(builder);                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message);
            }
        }

    }
}

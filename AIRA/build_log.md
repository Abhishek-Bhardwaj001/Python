### ================== A.I.R.A: AI Response Agent [V:1.0.0]: Baisc Voice Bot =====================

===================================Intitial Setup (22/08/26)=======================================
1. Started With writing controllers.voice_controller.py
   1. The Class will is the main entry code of our agent Handling text to speech functionality and conversions
   2. Call the class in main.py for the first initialization and validated the code conectivity with code placeholders.
   3. Setup Whisper class call
   4. Connect the Whisper text output with LLM call
   5. Connect the LLM response to tts for Vocals and Speach translation

2. Stt_services.py : Speach to text (Whisper)
3. Orchestrator_agent.py : AIRA Brain (Default: Groq)
4. tts_services.py : For Text to Speach Functionality (pyttsx3)


Learnings:
In a class You declare the variable local to method if it will change on its every call or declare it in the constructor if its going ot be static per object call.
If you import some variable from a file you dont have to declare it in the Class Constructor before using

Enhancements:
1. Make the Listener mic device Selection Dynamic (Let the User Pick)
2. Create a /settings through CLI or keep user freedom for changing Orchestrator brain mid session.
3. Add log variable for controlling print statements

Tech Debt:
1. Clean you Code and Write Doc String for every class and Function
### ================== A.I.R.A: AI Response Agent [V:1.0.0]: Basic Voice Bot =====================

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

Day 2 [23/08/26]:
Added Session Memory for AIRA: The bot looks to be working fine, however i got difficulty in connecting the Orchestrator class with LangGraph workflow specially the messages and memory State

Day 3 [24/08/26]
1. Switched to Elevenlabs for TTS Services.
2. Wired the Streamlit session history into the LangGraph backend.

Decision: who owns the conversation?
The checkpointer owns it, the UI only renders it.
- Workflow.run(query, thread_id, history=None). thread_id identifies the conversation.
- `history` is ONLY used to rehydrate a thread the checkpointer has never seen (fresh
  browser session, or the server restarting while the browser still holds state).
  Once the thread exists the arg is ignored, so nothing gets inserted twice.
- Cold-thread check is `app.get_state(cfg).values["messages"]`.
- CLI uses a fixed thread ("local-cli"), Streamlit uses a per-session uuid4.

Bugs found and fixed:
1. voice_controller: `def _get_agent_response(self,text,history=history)` -> NameError at
   import. Default args are evaluated when the class body runs, so `history` did not exist.
   The whole app was un-importable.
2. stream_app: VoiceController was built at module scope. Streamlit re-runs the script on
   every interaction, so each turn rebuilt Workflow() with a fresh MemorySaver and wiped
   the memory. THIS was the real reason memory never worked. Fixed with @st.cache_resource.
   It also reloaded the Whisper model every turn.
3. workflow.run accepted `history` and silently discarded it.
4. thread_id was hardcoded "default" -> every browser session shared one conversation.
5. stream_app: the assistant render block was indented under `if prompt:` instead of
   `if result:` -> NameError on `response` whenever result was None. Also dead code, since
   st.rerun() threw the render away.
6. helpers: st.audio(prompt.audio) consumes the buffer, so the following .read() returned
   empty. Needs seek(0) first.
7. stt_services: `sum(...)/len(segments)` with no empty check -> ZeroDivisionError on silence.
8. stt_services: whisper.transcribe() on a raw numpy array assumes 16 kHz MONO and does no
   resampling. Browser audio is 48 kHz and often stereo, so transcripts were garbled.
   Now downmixes to mono and resamples (scipy resample_poly, numpy.interp fallback).
   Note: Listen already resamples to 16 kHz, so the CLI path must NOT convert again.
9. nodes: returned `state.messages + [AIMessage]` while the state uses the add_messages
   reducer. Only the delta should be returned.
10. tts_services: reusing one pyttsx3 engine for save_to_file/runAndWait dies on Windows
    SAPI5 after the first call, which is exactly what a cached resource does. Engine is now
    built per call. Temp .wav files were also leaking, one per turn.
11. orchestrator_agent: unsupported engine printed and returned None -> AttributeError later.
    Now raises. Missing GROQ_API also raises instead of failing deep in the stack.
12. STT returning None flowed into HumanMessage(content=None). Guarded in the controller.

Added run_ui.py: `streamlit run` puts the SCRIPT's directory on sys.path, not the project
root, so `from app...` could never resolve when pointing it at stream_app.py directly.
Run the UI with `streamlit run run_ui.py`.

Learnings:
Default argument values are evaluated once, when the `def` line executes, not per call.
So `def f(self, x=x)` needs `x` to already exist in the enclosing scope.
Streamlit re-runs your whole script on every interaction. Anything expensive or stateful
(models, checkpointers, DB connections) must go behind @st.cache_resource or it is rebuilt
and its state lost on every turn.
When state uses a reducer like add_messages, a node returns only the NEW messages.

In a class You declare the variable local to method if it will change on its every call or declare it in the constructor if its going ot be static per object call.
If you import some variable from a file you dont have to declare it in the Class Constructor before using

Enhancements:
1. Make the Listener mic device Selection Dynamic (Let the User Pick)
2. Create a /settings through CLI or keep user freedom for changing Orchestrator brain mid session.
3. Add log variable for controlling print statements
4. MemorySaver is in-process, so a server restart still loses the thread (the rehydrate path
   covers the common case). Swap for SqliteSaver for real durability.
5. The prompt says no special characters since the text gets read aloud, but the model still
   emits **markdown bold**. Either strip markdown before TTS or tighten the prompt.
6. app/ui/streamlit/ shares a name with the streamlit package. Harmless today, but renaming
   it to app/ui/web/ would remove the footgun.

Tech Debt:
1. Clean you Code and Write Doc String for every class and Function
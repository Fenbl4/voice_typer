import threading
import time
import unittest
from pathlib import Path
import sys

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import voice_typer


class AudioPreparationTests(unittest.TestCase):
    def test_silence_is_rejected(self):
        silence = np.zeros(voice_typer.SAMPLE_RATE, dtype=np.int16)
        self.assertEqual(voice_typer._trim_pcm_silence(silence.tobytes()), b"")

    def test_edge_silence_is_trimmed_with_speech_padding(self):
        sample_rate = voice_typer.SAMPLE_RATE
        tone = (
            np.sin(np.arange(sample_rate // 2) * 2 * np.pi * 440 / sample_rate) * 3000
        ).astype(np.int16)
        padded = np.concatenate(
            [
                np.zeros(sample_rate // 2, dtype=np.int16),
                tone,
                np.zeros(sample_rate, dtype=np.int16),
            ]
        )

        trimmed = voice_typer._trim_pcm_silence(padded.tobytes())
        duration = len(trimmed) / 2 / sample_rate

        self.assertGreater(duration, 0.65)
        self.assertLess(duration, 0.90)


class TranscriptCleanupTests(unittest.TestCase):
    def test_known_groq_tail_is_removed_only_as_separate_final_sentence(self):
        source = "Основной текст. Продолжение следует..."
        self.assertEqual(voice_typer._clean_transcript(source, "Groq"), "Основной текст.")

    def test_local_text_is_not_semantically_changed(self):
        source = "Основной текст. Продолжение следует..."
        self.assertEqual(voice_typer._clean_transcript(source, "Local"), source)

    def test_words_inside_real_sentence_are_not_removed(self):
        source = "И дальше продолжение следует."
        self.assertEqual(voice_typer._clean_transcript(source, "Groq"), source)


class RecordingLifecycleTests(unittest.TestCase):
    def test_processing_waits_for_last_recording_frame(self):
        app = voice_typer.VoiceTyperApp.__new__(voice_typer.VoiceTyperApp)
        app._processing_lock = threading.Lock()
        app.audio_frames = [b"first"]
        app._target_hwnd = 17
        app._log = lambda *_: None
        captured = {}

        def finish_recording():
            time.sleep(0.03)
            app.audio_frames.append(b"last")

        app._record_thread = threading.Thread(target=finish_recording)
        app._record_thread.start()
        app._do_transcribe = lambda frames, hwnd: captured.update(frames=frames, hwnd=hwnd)

        app._process_and_paste(wait_for_recording=True)

        self.assertEqual(captured["frames"], [b"first", b"last"])
        self.assertEqual(captured["hwnd"], 17)


if __name__ == "__main__":
    unittest.main()

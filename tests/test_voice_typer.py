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
        app._do_transcribe = lambda frames, hwnd, live=None: captured.update(frames=frames, hwnd=hwnd)

        app._process_and_paste(wait_for_recording=True)

        self.assertEqual(captured["frames"], [b"first", b"last"])
        self.assertEqual(captured["hwnd"], 17)

    def test_processing_stops_live_worker_before_transcription(self):
        app = voice_typer.VoiceTyperApp.__new__(voice_typer.VoiceTyperApp)
        app._processing_lock = threading.Lock()
        app.audio_frames = [b"a", b"b"]
        app._target_hwnd = 5
        app._record_thread = None
        app._log = lambda *_: None
        captured = {}

        live = voice_typer._LiveTranscription(app.audio_frames)
        live.thread = threading.Thread(target=lambda: live.stop.wait(5))
        live.thread.start()
        app._do_transcribe = lambda frames, hwnd, live=None: captured.update(live=live)

        app._process_and_paste(wait_for_recording=True, live=live)

        self.assertIs(captured["live"], live)
        self.assertFalse(live.thread.is_alive())
        self.assertTrue(live.usable())


def _frames_from_pcm(pcm: bytes) -> list[bytes]:
    step = voice_typer.CHUNK_SIZE * 2
    return [pcm[i:i + step] for i in range(0, len(pcm), step)]


class LiveTranscriptionTests(unittest.TestCase):
    """Живое распознавание во время записи: закрытые отрезки уходят в текст, открытый хвост ждёт отпускания."""

    def _make_app(self, segments_by_call, texts):
        app = voice_typer.VoiceTyperApp.__new__(voice_typer.VoiceTyperApp)
        app._log = lambda *_: None
        calls = []

        def speech_segments(waveform):
            calls.append(waveform.size)
            return segments_by_call.pop(0)

        class FakeModel:
            def __init__(self):
                self.seen = []

            def recognize(self, waveform, sample_rate=16000):
                self.seen.append(waveform.size)
                return texts.pop(0)

        app._local_speech_segments = speech_segments
        app._local_model = FakeModel()
        return app, calls

    def test_short_recording_is_not_split(self):
        # Записи короче порога живой нарезки распознаются целиком, как раньше.
        sample_rate = voice_typer.SAMPLE_RATE
        frames = _frames_from_pcm(np.zeros(sample_rate * 18, dtype=np.int16).tobytes())
        app, calls = self._make_app([[]], [])
        live = voice_typer._LiveTranscription(frames)

        app._live_transcribe_step(live)

        self.assertEqual(calls, [])
        self.assertEqual(live.consumed_frames, 0)

    def test_short_pause_inside_phrase_does_not_cut_it(self):
        # Заминка короче 1 с («бу... будет») склеивается с продолжением, пауза от 1 с режет.
        sample_rate = voice_typer.SAMPLE_RATE
        merged = voice_typer._merge_close_segments(
            [(0, sample_rate * 3), (sample_rate * 3 + sample_rate // 2, sample_rate * 6),
             (sample_rate * 8, sample_rate * 9)],
            sample_rate, sample_rate * 22,
        )
        self.assertEqual(merged, [(0, sample_rate * 6), (sample_rate * 8, sample_rate * 9)])

    def test_merge_respects_max_segment_length(self):
        sample_rate = voice_typer.SAMPLE_RATE
        merged = voice_typer._merge_close_segments(
            [(0, sample_rate * 21), (sample_rate * 21 + 100, sample_rate * 30)],
            sample_rate, sample_rate * 22,
        )
        self.assertEqual(merged, [(0, sample_rate * 21), (sample_rate * 21 + 100, sample_rate * 30)])

    def test_closed_segments_are_consumed_and_open_tail_is_kept(self):
        sample_rate = voice_typer.SAMPLE_RATE
        total = sample_rate * 24
        frames = _frames_from_pcm(np.zeros(total, dtype=np.int16).tobytes())
        closed_end = sample_rate * 5
        open_start = sample_rate * 8
        app, calls = self._make_app(
            [[(sample_rate, closed_end), (open_start, total)]],
            ["первый отрезок"],
        )
        live = voice_typer._LiveTranscription(frames)

        app._live_transcribe_step(live)

        self.assertEqual(calls, [total])
        self.assertEqual(live.texts, ["первый отрезок"])
        self.assertEqual(app._local_model.seen, [closed_end - sample_rate])
        # Указатель округляется вниз до границы блока: ни один отсчёт речи не теряется,
        # а в хвост может попасть меньше одного блока (64 мс) тишины из запаса отрезка.
        consumed_samples = live.consumed_frames * voice_typer.CHUNK_SIZE
        self.assertGreater(consumed_samples, closed_end - voice_typer.CHUNK_SIZE)
        self.assertLessEqual(consumed_samples, closed_end)
        self.assertLess(consumed_samples, open_start)

    def test_next_step_starts_after_consumed_frames(self):
        sample_rate = voice_typer.SAMPLE_RATE
        first_total = sample_rate * 24
        frames = _frames_from_pcm(np.zeros(first_total, dtype=np.int16).tobytes())
        app, calls = self._make_app(
            [[(0, sample_rate * 4)], [(sample_rate, sample_rate * 3)]],
            ["раз", "два"],
        )
        live = voice_typer._LiveTranscription(frames)
        app._live_transcribe_step(live)
        first_consumed = live.consumed_frames

        frames.extend(_frames_from_pcm(np.zeros(sample_rate * 10, dtype=np.int16).tobytes()))
        app._live_transcribe_step(live)

        self.assertEqual(live.texts, ["раз", "два"])
        self.assertEqual(calls[1], sum(len(f) for f in frames[first_consumed:]) // 2)
        self.assertGreater(live.consumed_frames, first_consumed)

    def test_worker_failure_marks_live_unusable(self):
        sample_rate = voice_typer.SAMPLE_RATE
        frames = _frames_from_pcm(np.zeros(sample_rate * 24, dtype=np.int16).tobytes())
        app = voice_typer.VoiceTyperApp.__new__(voice_typer.VoiceTyperApp)
        app._log = lambda *_: None

        def broken(_waveform):
            raise RuntimeError("vad broke")

        app._local_speech_segments = broken
        live = voice_typer._LiveTranscription(frames)
        live.stop = type("Once", (), {"wait": staticmethod(lambda _t: False), "is_set": staticmethod(lambda: False)})()

        app._live_transcribe_loop(live)

        self.assertTrue(live.failed)
        self.assertFalse(live.usable())


if __name__ == "__main__":
    unittest.main()

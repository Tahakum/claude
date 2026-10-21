"""Synthesize a 30s Saudi/Khaleeji-flavored music bed timed to the promo scenes.

Elements: khaleeji dum/tak percussion (khabiti-like), daff-ish shaker, claps,
Karplus-Strong plucked 'oud' in maqam Hijaz on D, sub bass, low drone.
"""
import numpy as np
from scipy.signal import butter, sosfilt
from scipy.io import wavfile

SR = 48000
DUR = 30.0
rng = np.random.default_rng(11)
track = np.zeros(int(DUR * SR))

def bp(x, lo, hi):
    return sosfilt(butter(4, [lo, hi], btype="band", fs=SR, output="sos"), x)

def lp(x, fc):
    return sosfilt(butter(4, fc, btype="low", fs=SR, output="sos"), x)

def env(n, a, d):
    """attack/decay envelope"""
    e = np.ones(n)
    na = max(1, int(a * SR))
    e[:na] = np.linspace(0, 1, na)
    t = np.arange(n) / SR
    e *= np.exp(-t * d)
    return e

def place(sound, at, vol=1.0):
    i = int(at * SR)
    if i >= len(track):
        return
    j = min(len(track), i + len(sound))
    track[i:j] += sound[: j - i] * vol

# ---------- instruments ----------
def dum(vol=1.0):
    """Deep tabl/daff dum."""
    n = int(0.35 * SR)
    t = np.arange(n) / SR
    f = 120 * np.exp(-t * 22) + 58
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR)
    skin = bp(rng.standard_normal(n), 120, 500) * 0.25
    return (tone + skin) * env(n, 0.002, 14) * vol

def tak(vol=1.0):
    """Sharp rim/slap."""
    n = int(0.09 * SR)
    x = bp(rng.standard_normal(n), 1800, 6500) * env(n, 0.001, 70)
    x += bp(rng.standard_normal(n), 700, 1400) * env(n, 0.001, 55) * 0.4
    return x * vol

def shk(vol=1.0):
    """Daff jingle / shaker tick."""
    n = int(0.05 * SR)
    return bp(rng.standard_normal(n), 5000, 11000) * env(n, 0.001, 110) * vol

def clap(vol=1.0):
    n = int(0.25 * SR)
    x = np.zeros(n)
    for d in [0, 0.012, 0.025]:
        i = int(d * SR)
        m = bp(rng.standard_normal(n - i), 900, 3200) * env(n - i, 0.001, 38)
        x[i:] += m
    return x * vol

def pluck(freq, dur=0.5, vol=1.0, bright=0.55):
    """Karplus-Strong oud-ish pluck, doubled slightly detuned."""
    out = np.zeros(int(dur * SR))
    for det in [1.0, 1.003]:
        p = int(SR / (freq * det))
        buf = (rng.random(p) * 2 - 1) * (1 - bright) + bright * np.sin(np.linspace(0, 2 * np.pi, p))
        n = len(out)
        y = np.zeros(n)
        idx = 0
        for i in range(n):
            y[i] = buf[idx]
            buf[idx] = 0.5 * (buf[idx] + buf[(idx + 1) % p]) * 0.996
            idx = (idx + 1) % p
        out += y * 0.5
    out = lp(out, 4200)
    t = np.arange(len(out)) / SR
    return out * np.exp(-t * 2.2) * vol

def sub(freq, dur, vol=1.0):
    n = int(dur * SR)
    t = np.arange(n) / SR
    x = np.sin(2 * np.pi * freq * t) + 0.3 * np.sin(2 * np.pi * freq * 2 * t)
    e = np.minimum(1, t / 0.02) * np.exp(-t * 3.5)
    return x * e * vol

def drone(freq, dur, vol=1.0):
    n = int(dur * SR)
    t = np.arange(n) / SR
    x = np.sin(2 * np.pi * freq * t) * 0.6 + np.sin(2 * np.pi * freq * 1.5 * t) * 0.2
    x += bp(rng.standard_normal(n), freq * 0.9, freq * 3) * 0.12
    fade = np.minimum(1, t / 1.2) * np.minimum(1, (dur - t) / 1.2)
    return lp(x, 900) * fade * vol

# Hijaz on D (D Eb F# G A Bb C)
D3, Eb3, Gb3, G3, A3, Bb3, C4, D4 = 146.83, 155.56, 185.0, 196.0, 220.0, 233.08, 261.63, 293.66
D2 = 73.42

B = 60 / 102  # beat @102bpm

def groove(start, bars, intensity=1.0, with_clap=True):
    """2-beat khabiti-like bar: DUM . tak DUM DUM . tak ."""
    for b in range(bars):
        t0 = start + b * 2 * B
        place(dum(0.95), t0, intensity)
        place(tak(0.5), t0 + 0.75 * B, intensity)
        place(dum(0.8), t0 + 1.0 * B, intensity)
        place(dum(0.55), t0 + 1.25 * B, intensity * 0.8)
        place(tak(0.6), t0 + 1.75 * B, intensity)
        for e in range(4):
            place(shk(0.3), t0 + e * 0.5 * B + 0.25 * B, intensity * 0.7)
        if with_clap and b % 2 == 1:
            place(clap(0.5), t0 + 1.0 * B, intensity)
        place(sub(D2, 0.5, 0.7), t0, intensity)
        place(sub(D2, 0.3, 0.45), t0 + 1.0 * B, intensity * 0.8)

# ---------- arrangement ----------
# S1+S2 (0-6.5): dark drone + sparse heartbeat dums on text beats
place(drone(D2, 6.6, 0.5), 0)
for at in [0.15, 1.3, 2.65, 3.07, 3.95, 4.37, 5.25, 5.67]:
    place(dum(0.8), at, 0.8)
place(shk(0.4), 2.0)
place(shk(0.4), 4.7)

# silence gap 6.5-6.9 (nothing)

# S3 reveal (6.9-9.5): daff roll into downbeat + oud hijaz signature
roll_t = 7.1
for i in range(10):
    place(tak(0.35 + i * 0.05), roll_t + i * 0.045, 0.7)
place(dum(1.0), 7.55, 1.1)
place(clap(0.7), 7.55)
place(sub(D2, 0.9, 1.0), 7.55)
# signature hijaz ascent: D Eb F# A -> D4
for note, dt, dur in [(D3, 0.0, 0.3), (Eb3, 0.18, 0.3), (Gb3, 0.36, 0.3), (A3, 0.54, 0.45), (D4, 0.9, 1.4)]:
    place(pluck(note, dur + 0.3, 0.85), 7.8 + dt)
place(dum(0.9), 8.75, 0.9)
place(tak(0.6), 9.05, 0.8)
place(dum(0.9), 9.15, 0.9)

# S4 hakeem (9.5-19): groove + oud ostinato (keep mid sparse under captions)
groove(9.6, 8, intensity=0.85)
ost = [(D3, 0), (A3, 0.5), (G3, 1.0), (Gb3, 1.5), (D3, 2.0), (A3, 2.5), (Bb3, 3.0), (A3, 3.5)]
for rep in range(2):
    for note, beats in ost:
        place(pluck(note, 0.5, 0.5), 9.6 + rep * 4 * 2 * B / 2 + beats * B)
# answer moment accent (15.3)
for note, dt in [(D4, 0), (C4, 0.22), (Bb3, 0.44), (A3, 0.66)]:
    place(pluck(note, 0.8, 0.7), 15.3 + dt)

# S5 services (19-25): groove up + phrase per cut
groove(19.0, 5, intensity=1.0)
for cut in [19.0, 21.0, 23.0]:
    place(dum(1.0), cut, 1.05)
    place(clap(0.6), cut + 0.07)
for note, dt in [(A3, 19.3), (G3, 19.9), (D4, 21.3), (C4, 21.9), (A3, 23.3), (D4, 23.9)]:
    place(pluck(note, 0.7, 0.6), dt)

# S6 CTA (25-30): climax then cadence
place(dum(1.1), 25.15, 1.15)
place(clap(0.8), 25.15)
place(sub(D2, 1.2, 1.0), 25.15)
groove(25.3, 3, intensity=0.9, with_clap=False)
# final cadence: DUM DUM TAK + low oud strum, then tail
place(dum(1.0), 28.55, 1.0)
place(dum(1.0), 28.85, 1.0)
place(tak(0.8), 29.1, 0.9)
for note, dt in [(D3, 0), (A3, 0.02), (D4, 0.05)]:
    place(pluck(note, 1.6, 0.8), 29.15 + dt)
place(sub(D2, 1.4, 0.9), 29.15)

# ---------- master ----------
track[-int(0.6 * SR):] *= np.linspace(1, 0, int(0.6 * SR))
track = np.tanh(track * 0.9)
track = track / np.max(np.abs(track)) * 0.8
stereo = np.stack([track, track], axis=1)
wavfile.write("../audio/music_khaleeji.wav", SR, (stereo * 32767).astype(np.int16))
print("music_khaleeji.wav written")

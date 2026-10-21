"""Ardah-flavored majestic Saudi percussion bed, 30s, timed to the promo scenes.

Big ensemble drums (takhmir/zeer), tar taks, row claps, light reverb,
sparse oud accents in Hijaz-D. Simple, familiar, heavy.
"""
import numpy as np
from scipy.signal import butter, sosfilt, fftconvolve
from scipy.io import wavfile

SR = 48000
DUR = 30.0
rng = np.random.default_rng(23)
track = np.zeros(int(DUR * SR))

def bp(x, lo, hi):
    return sosfilt(butter(4, [lo, hi], btype="band", fs=SR, output="sos"), x)

def lp(x, fc):
    return sosfilt(butter(4, fc, btype="low", fs=SR, output="sos"), x)

def env(n, a, d):
    e = np.ones(n)
    na = max(1, int(a * SR))
    e[:na] = np.linspace(0, 1, na)
    e *= np.exp(-np.arange(n) / SR * d)
    return e

def place(sound, at, vol=1.0):
    i = int(at * SR)
    if i >= len(track) or i < 0:
        return
    j = min(len(track), i + len(sound))
    track[i:j] += sound[: j - i] * vol

# ---------- single hits ----------
def zeer_hit(pitch=1.0):
    """One deep ardah bass drum stroke."""
    n = int(0.5 * SR)
    t = np.arange(n) / SR
    f = (105 * np.exp(-t * 16) + 48) * pitch
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR)
    skin = bp(rng.standard_normal(n), 90, 420) * 0.35
    thud = lp(rng.standard_normal(n), 160) * 0.5
    return (tone * 1.1 + skin + thud) * env(n, 0.003, 11)

def tar_tak():
    """Tar frame-drum rim tak."""
    n = int(0.12 * SR)
    x = bp(rng.standard_normal(n), 1400, 5200) * env(n, 0.001, 55)
    x += bp(rng.standard_normal(n), 500, 1100) * env(n, 0.001, 45) * 0.5
    return x

def row_clap():
    """Group clap (several hands, slightly spread)."""
    n = int(0.3 * SR)
    x = np.zeros(n)
    for d in (rng.random(5) * 0.03):
        i = int(d * SR)
        m = bp(rng.standard_normal(n - i), 800, 3400) * env(n - i, 0.001, 34)
        x[i:] += m * (0.6 + 0.4 * rng.random())
    return x * 0.5

def ensemble(hit_fn, layers=4, spread=0.025):
    """Stack slightly detuned/delayed copies -> crowd of drummers."""
    base = hit_fn()
    n = len(base) + int(spread * SR) + 1
    out = np.zeros(n)
    out[: len(base)] += base
    for _ in range(layers - 1):
        h = hit_fn()
        i = int(rng.random() * spread * SR)
        out[i : i + len(h)] += h * (0.55 + 0.35 * rng.random())
    return out / layers * 2.2

def pluck(freq, dur=0.5, vol=1.0):
    out = np.zeros(int(dur * SR))
    for det in [1.0, 1.004]:
        p = int(SR / (freq * det))
        buf = (rng.random(p) * 2 - 1) * 0.5 + 0.5 * np.sin(np.linspace(0, 2 * np.pi, p))
        y = np.zeros(len(out))
        idx = 0
        for i in range(len(out)):
            y[i] = buf[idx]
            buf[idx] = 0.5 * (buf[idx] + buf[(idx + 1) % p]) * 0.996
            idx = (idx + 1) % p
        out += y * 0.5
    out = lp(out, 3800)
    return out * np.exp(-np.arange(len(out)) / SR * 2.4) * vol

def sub(freq, dur, vol=1.0):
    n = int(dur * SR)
    t = np.arange(n) / SR
    x = np.sin(2 * np.pi * freq * t)
    return x * np.minimum(1, t / 0.02) * np.exp(-t * 3.2) * vol

def drone(freq, dur, vol=1.0):
    n = int(dur * SR)
    t = np.arange(n) / SR
    x = np.sin(2 * np.pi * freq * t) * 0.6 + np.sin(2 * np.pi * freq * 1.5 * t) * 0.18
    x += bp(rng.standard_normal(n), freq, freq * 3) * 0.1
    fade = np.minimum(1, t / 1.2) * np.minimum(1, (dur - t) / 1.2)
    return lp(x, 800) * fade * vol

D3, Eb3, Gb3, G3, A3, D4 = 146.83, 155.56, 185.0, 196.0, 220.0, 293.66
D2 = 73.42
B = 60 / 88  # 88 bpm, ardah-slow

def ardah_bar(t0, intensity=1.0, claps=True):
    """One 2-beat ardah bar: DUM . tak | DUM-DUM . tak, group clap on beat 2."""
    place(ensemble(zeer_hit), t0, 1.0 * intensity)
    place(sub(D2, 0.45, 0.6), t0, intensity)
    place(ensemble(tar_tak, 3, 0.015), t0 + 0.75 * B, 0.8 * intensity)
    place(ensemble(lambda: zeer_hit(1.06), 3), t0 + 1.0 * B, 0.75 * intensity)
    place(ensemble(lambda: zeer_hit(0.97), 3), t0 + 1.25 * B, 0.6 * intensity)
    place(ensemble(tar_tak, 3, 0.015), t0 + 1.75 * B, 0.85 * intensity)
    if claps:
        place(row_clap(), t0 + 1.0 * B, 0.9 * intensity)

# ---------- arrangement ----------
# S1+S2 (0-6.5): dark drone + lone heavy zeer on text beats
place(drone(D2, 6.6, 0.45), 0)
for at in [0.15, 1.3, 2.65, 3.07, 3.95, 4.37, 5.25, 5.67]:
    place(ensemble(zeer_hit, 3), at, 0.75)

# S3 reveal (6.9-9.5): tar roll -> first full bar + oud signature
for i in range(12):
    place(tar_tak(), 7.05 + i * 0.042, 0.45 + i * 0.045)
place(ensemble(zeer_hit, 5), 7.55, 1.15)
place(row_clap(), 7.55, 1.0)
place(sub(D2, 0.9, 0.95), 7.55)
for note, dt, dr in [(D3, 0.0, 0.35), (Gb3, 0.22, 0.35), (A3, 0.44, 0.5), (D4, 0.8, 1.3)]:
    place(pluck(note, dr + 0.3, 0.6), 7.85 + dt)
ardah_bar(8.6, 0.8)

# S4 hakeem (9.6-19): steady ardah, slightly reserved under captions
tt = 9.6
while tt < 18.6:
    ardah_bar(tt, 0.8)
    tt += 2 * B
place(pluck(A3, 0.6, 0.4), 11.9)
place(pluck(D4, 0.9, 0.5), 15.35)
place(pluck(A3, 0.7, 0.4), 15.65)

# S5 services (19-25): full power, hit on each cut
tt = 19.0
while tt < 24.6:
    ardah_bar(tt, 1.0)
    tt += 2 * B
for cut in [19.0, 21.0, 23.0]:
    place(ensemble(zeer_hit, 5), cut, 1.1)
    place(row_clap(), cut + 0.06, 1.0)
place(pluck(D4, 0.8, 0.5), 21.3)
place(pluck(A3, 0.8, 0.5), 23.3)

# S6 CTA (25-30): climax, wink ting, cadence
place(ensemble(zeer_hit, 6), 25.15, 1.2)
place(row_clap(), 25.15, 1.1)
place(sub(D2, 1.1, 1.0), 25.15)
ardah_bar(25.6, 0.9)
ardah_bar(25.6 + 2 * B, 0.85)
# wink ting (28.45): high oud harmonic + tiny tak
place(pluck(D4 * 2, 0.5, 0.5), 28.45)
place(tar_tak(), 28.47, 0.5)
# final cadence
place(ensemble(zeer_hit, 5), 28.95, 1.05)
place(ensemble(zeer_hit, 6), 29.25, 1.15)
place(row_clap(), 29.25, 1.0)
for note, dt in [(D3, 0), (A3, 0.03), (D4, 0.06)]:
    place(pluck(note, 1.4, 0.7), 29.3 + dt)
place(sub(D2, 1.2, 0.95), 29.3)

# ---------- light hall reverb ----------
ir_n = int(0.38 * SR)
ir = (rng.standard_normal(ir_n) * np.exp(-np.arange(ir_n) / SR * 14))
ir = bp(ir, 300, 6000)
wet = fftconvolve(track, ir)[: len(track)] * 0.055
track = track + wet

# ---------- master ----------
track[-int(0.5 * SR):] *= np.linspace(1, 0, int(0.5 * SR))
track = np.tanh(track * 0.85)
track = track / np.max(np.abs(track)) * 0.8
stereo = np.stack([track, track], axis=1)
wavfile.write("../audio/music_ardah.wav", SR, (stereo * 32767).astype(np.int16))
print("music_ardah.wav written")

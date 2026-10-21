"""v4: Upbeat Gulf-pop ad groove (the actual Saudi app-ad sound: soft kick,
snaps, light hats, warm bass, subtle pads) + oud only at reveal/outro +
the approved ardah stamp on Hakeem's wink.
"""
import numpy as np
from scipy.signal import butter, sosfilt, fftconvolve
from scipy.io import wavfile

SR = 48000
DUR = 30.0
rng = np.random.default_rng(47)
track = np.zeros(int(DUR * SR))

def bp(x, lo, hi):
    return sosfilt(butter(4, [lo, hi], btype="band", fs=SR, output="sos"), x)

def lp(x, fc):
    return sosfilt(butter(4, fc, btype="low", fs=SR, output="sos"), x)

def env(n, a, d):
    e = np.ones(n); na = max(1, int(a * SR))
    e[:na] = np.linspace(0, 1, na)
    return e * np.exp(-np.arange(n) / SR * d)

def place(sound, at, vol=1.0):
    i = int(at * SR)
    if i >= len(track) or i < 0: return
    j = min(len(track), i + len(sound))
    track[i:j] += sound[: j - i] * vol

# ---------- kit ----------
def kick():
    n = int(0.28 * SR); t = np.arange(n) / SR
    f = 130 * np.exp(-t * 30) + 52
    x = np.sin(2 * np.pi * np.cumsum(f) / SR)
    click = bp(rng.standard_normal(n), 900, 2600) * env(n, 0.001, 150) * 0.25
    return (x * env(n, 0.002, 16) + click)

def snap():
    n = int(0.2 * SR)
    x = np.zeros(n)
    for d in (rng.random(2) * 0.008):
        i = int(d * SR)
        x[i:] += bp(rng.standard_normal(n - i), 1500, 4500) * env(n - i, 0.001, 48)
    return x * 0.6

def hat(open_=False):
    n = int((0.12 if open_ else 0.045) * SR)
    return bp(rng.standard_normal(n), 7000, 13000) * env(n, 0.001, 40 if open_ else 130) * 0.4

def daff_touch():
    n = int(0.16 * SR); t = np.arange(n) / SR
    f = 90 * np.exp(-t * 26) + 70
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * env(n, 0.003, 24) * 0.5

def bass_note(freq, dur, vol=1.0):
    n = int(dur * SR); t = np.arange(n) / SR
    x = np.sin(2 * np.pi * freq * t) * 0.75
    ph = 2 * np.pi * freq * t
    x += lp(2 * ((ph / (2 * np.pi)) % 1) - 1, 420) * 0.35
    e = np.minimum(1, t / 0.015) * np.exp(-t * 2.2)
    return x * e * vol

def pad(freqs, dur, vol=1.0, warmth=1300):
    n = int(dur * SR); t = np.arange(n) / SR
    x = np.zeros(n)
    for f0 in freqs:
        for det in [0.997, 1.0, 1.004]:
            ph = 2 * np.pi * f0 * det * t
            x += (2 * ((ph / (2 * np.pi)) % 1) - 1) * 0.07
    x = lp(x, warmth)
    swell = np.minimum(1, t / 1.2) * np.minimum(1, np.maximum(0.0001, (dur - t)) / 1.0)
    return x * swell * vol

def pluck(freq, dur=0.5, vol=1.0):
    out = np.zeros(int(dur * SR))
    for det in [1.0, 1.004]:
        p = int(SR / (freq * det))
        buf = (rng.random(p) * 2 - 1) * 0.5 + 0.5 * np.sin(np.linspace(0, 2 * np.pi, p))
        y = np.zeros(len(out)); idx = 0
        for i in range(len(out)):
            y[i] = buf[idx]
            buf[idx] = 0.5 * (buf[idx] + buf[(idx + 1) % p]) * 0.996
            idx = (idx + 1) % p
        out += y * 0.5
    return lp(out, 3600) * np.exp(-np.arange(len(out)) / SR * 2.6) * vol

def zeer_hit(pitch=1.0):
    n = int(0.5 * SR); t = np.arange(n) / SR
    f = (105 * np.exp(-t * 16) + 48) * pitch
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR)
    return (tone * 1.1 + bp(rng.standard_normal(n), 90, 420) * 0.35 + lp(rng.standard_normal(n), 160) * 0.5) * env(n, 0.003, 11)

def tar_tak():
    n = int(0.12 * SR)
    return bp(rng.standard_normal(n), 1400, 5200) * env(n, 0.001, 55) + bp(rng.standard_normal(n), 500, 1100) * env(n, 0.001, 45) * 0.5

def ensemble(hit_fn, layers=4, spread=0.025):
    base = hit_fn()
    out = np.zeros(len(base) + int(spread * SR) + 1)
    out[: len(base)] += base
    for _ in range(layers - 1):
        h = hit_fn(); i = int(rng.random() * spread * SR)
        out[i : i + len(h)] += h * (0.55 + 0.35 * rng.random())
    return out / layers * 2.2

def row_clap():
    n = int(0.3 * SR); x = np.zeros(n)
    for d in (rng.random(5) * 0.03):
        i = int(d * SR)
        x[i:] += bp(rng.standard_normal(n - i), 800, 3400) * env(n - i, 0.001, 34) * (0.6 + 0.4 * rng.random())
    return x * 0.5

D2, A2, Bb1, C2 = 73.42, 110.0, 58.27, 65.41
D3, Gb3, G3, A3, Bb3, C4, D4 = 146.83, 185.0, 196.0, 220.0, 233.08, 261.63, 293.66
B = 60 / 100  # 100 bpm

def groove_bar(t0, intensity=1.0, hats=True, daff=False):
    """4-beat pop bar: kick 1&3, snap 2&4, hats 8ths w/ light swing."""
    place(kick(), t0, 0.95 * intensity)
    place(kick(), t0 + 2 * B, 0.85 * intensity)
    place(kick(), t0 + 2.75 * B, 0.45 * intensity)
    place(snap(), t0 + 1 * B, 0.9 * intensity)
    place(snap(), t0 + 3 * B, 0.95 * intensity)
    if hats:
        for e in range(8):
            sw = 0.035 * B if e % 2 == 1 else 0
            place(hat(e == 7), t0 + e * 0.5 * B + sw, (0.5 if e % 2 else 0.7) * intensity)
    if daff:
        place(daff_touch(), t0 + 1.5 * B, 0.6 * intensity)
        place(daff_touch(), t0 + 3.5 * B, 0.45 * intensity)

bassline = [(D2, 0, 0.7), (D2, 0.75, 0.3), (A2, 1.5, 0.45), (D2, 2.0, 0.7), (C2, 3.0, 0.5), (Bb1, 3.5, 0.5)]
def bass_bar(t0, root_shift=None, vol=0.6):
    for f, beat, dr in bassline:
        ff = f if root_shift is None else f * root_shift
        place(bass_note(ff, dr + 0.15, vol), t0 + beat * B)

# ---------- arrangement ----------
# S1+S2 (0-6.5): tense pulse — muted kick heartbeat + dark pad, accel ticks
place(pad([D2, A2], 6.8, 0.4, warmth=600), 0)
for at in [0.15, 1.3, 2.65, 3.07, 3.95, 4.37, 5.25, 5.67]:
    place(kick(), at, 0.8)
tt = 2.7; step = 0.5
while tt < 6.3:
    place(hat(), tt, 0.4); step = max(0.22, step * 0.9); tt += step

# S3 reveal (6.9-9.5): light opens -> groove starts softly + oud signature
place(pad([D3, A3, D4], 2.8, 0.5, warmth=1900), 6.95)
place(kick(), 7.55, 1.0); place(snap(), 7.58, 0.7)
place(bass_note(D2, 1.0, 0.8), 7.55)
for note, dt, dr in [(D3, 0.0, 0.3), (Gb3, 0.18, 0.3), (A3, 0.36, 0.45), (D4, 0.7, 1.1)]:
    place(pluck(note, dr + 0.3, 0.5), 7.9 + dt)
groove_bar(8.3, 0.6, hats=True)
bass_bar(8.3, vol=0.45)

# S4 (9.5-19): full groove, pads cycling, daff touches for identity
pads4 = [[D3, A3], [Bb1 * 4, G3], [C2 * 4, G3], [D3, A3]]
tt = 9.5; ci = 0
while tt < 18.7:
    groove_bar(tt, 0.85, daff=True)
    bass_bar(tt, vol=0.55)
    place(pad(pads4[ci % 4], 4 * B + 0.3, 0.3, warmth=1500), tt)
    tt += 4 * B; ci += 1
place(pluck(A3, 0.5, 0.3), 11.9)
place(pluck(D4, 0.7, 0.4), 15.35)

# S5 (19-25): lift — brighter pads, open hats, extra snap energy
tt = 19.0; ci = 0
while tt < 24.7:
    groove_bar(tt, 1.0, daff=True)
    bass_bar(tt, vol=0.65)
    place(pad([f * 2 for f in pads4[ci % 4]], 4 * B + 0.3, 0.22, warmth=2200), tt)
    place(pad(pads4[ci % 4], 4 * B + 0.3, 0.3, warmth=1600), tt)
    tt += 4 * B; ci += 1
for cut in [19.0, 21.0, 23.0]:
    place(kick(), cut, 1.05); place(snap(), cut + 0.03, 0.8)
place(pluck(D4, 0.7, 0.4), 21.3)
place(pluck(A3, 0.7, 0.35), 23.3)

# S6 (25-28.3): climax groove
place(kick(), 25.15, 1.1); place(snap(), 25.18, 0.9); place(bass_note(D2, 1.2, 0.85), 25.15)
place(pad([D3, A3, D4], 3.4, 0.45, warmth=2000), 25.2)
groove_bar(25.5, 0.9, daff=True)
bass_bar(25.5, vol=0.6)
groove_bar(25.5 + 4 * B, 0.75)

# WINK OUTRO (28.3-30): stop -> ting -> approved ardah heritage stamp
place(pluck(D4 * 2, 0.5, 0.5), 28.45)
place(tar_tak(), 28.47, 0.5)
place(ensemble(zeer_hit, 5), 28.95, 1.05)
place(ensemble(lambda: zeer_hit(1.05), 6), 29.25, 1.15)
place(row_clap(), 29.25, 1.0)
place(ensemble(tar_tak, 3, 0.015), 29.5, 0.8)
for note, dt in [(D3, 0), (A3, 0.03), (D4, 0.06)]:
    place(pluck(note, 1.3, 0.7), 29.3 + dt)
place(bass_note(D2, 1.1, 0.9), 29.25)

# ---------- glue ----------
ir_n = int(0.25 * SR)
ir = bp(rng.standard_normal(ir_n) * np.exp(-np.arange(ir_n) / SR * 22), 400, 6000)
track = track + fftconvolve(track, ir)[: len(track)] * 0.035
track[-int(0.5 * SR):] *= np.linspace(1, 0, int(0.5 * SR))
track = np.tanh(track * 0.92)
track = track / np.max(np.abs(track)) * 0.8
stereo = np.stack([track, track], axis=1)
wavfile.write("../audio/music_v4.wav", SR, (stereo * 32767).astype(np.int16))
print("music_v4.wav written")

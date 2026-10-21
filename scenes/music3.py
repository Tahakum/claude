"""v3: Modern-majlis Saudi bed — warm string pads + soft samri daff groove +
light oud, modern sub bass. Ardah ensemble appears ONLY at the Hakeem wink outro.
"""
import numpy as np
from scipy.signal import butter, sosfilt, fftconvolve
from scipy.io import wavfile

SR = 48000
DUR = 30.0
rng = np.random.default_rng(31)
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
    if i >= len(track) or i < 0: return
    j = min(len(track), i + len(sound))
    track[i:j] += sound[: j - i] * vol

# ---------- instruments ----------
def daff_dum():
    """Soft frame-drum dum (warm, not martial)."""
    n = int(0.3 * SR)
    t = np.arange(n) / SR
    f = 95 * np.exp(-t * 20) + 62
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR)
    skin = bp(rng.standard_normal(n), 150, 500) * 0.2
    return (tone + skin) * env(n, 0.004, 16)

def tik():
    n = int(0.07 * SR)
    return bp(rng.standard_normal(n), 2200, 7000) * env(n, 0.001, 90) * 0.7

def riqq():
    """Tambourine jingle shimmer."""
    n = int(0.12 * SR)
    return bp(rng.standard_normal(n), 6000, 12000) * env(n, 0.001, 55) * 0.5

def soft_clap():
    n = int(0.22 * SR)
    x = np.zeros(n)
    for d in (rng.random(3) * 0.018):
        i = int(d * SR)
        x[i:] += bp(rng.standard_normal(n - i), 1000, 3000) * env(n - i, 0.001, 40)
    return x * 0.4

def zeer_hit(pitch=1.0):
    n = int(0.5 * SR)
    t = np.arange(n) / SR
    f = (105 * np.exp(-t * 16) + 48) * pitch
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR)
    skin = bp(rng.standard_normal(n), 90, 420) * 0.35
    return (tone * 1.1 + skin + lp(rng.standard_normal(n), 160) * 0.5) * env(n, 0.003, 11)

def tar_tak():
    n = int(0.12 * SR)
    x = bp(rng.standard_normal(n), 1400, 5200) * env(n, 0.001, 55)
    return x + bp(rng.standard_normal(n), 500, 1100) * env(n, 0.001, 45) * 0.5

def ensemble(hit_fn, layers=4, spread=0.025):
    base = hit_fn()
    out = np.zeros(len(base) + int(spread * SR) + 1)
    out[: len(base)] += base
    for _ in range(layers - 1):
        h = hit_fn()
        i = int(rng.random() * spread * SR)
        out[i : i + len(h)] += h * (0.55 + 0.35 * rng.random())
    return out / layers * 2.2

def row_clap():
    n = int(0.3 * SR)
    x = np.zeros(n)
    for d in (rng.random(5) * 0.03):
        i = int(d * SR)
        x[i:] += bp(rng.standard_normal(n - i), 800, 3400) * env(n - i, 0.001, 34) * (0.6 + 0.4 * rng.random())
    return x * 0.5

def pad(freqs, dur, vol=1.0, warmth=1100):
    """Warm detuned-saw string pad with slow swell."""
    n = int(dur * SR)
    t = np.arange(n) / SR
    x = np.zeros(n)
    for f0 in freqs:
        for det in [0.996, 1.0, 1.005]:
            ph = 2 * np.pi * f0 * det * t + 0.1 * np.sin(2 * np.pi * 0.13 * t)
            saw = 2 * ((ph / (2 * np.pi)) % 1) - 1
            x += saw * 0.08
    x = lp(x, warmth)
    swell = np.minimum(1, t / 1.6) * np.minimum(1, np.maximum(0.0001,(dur - t)) / 1.4)
    breath = 1 + 0.12 * np.sin(2 * np.pi * 0.21 * t)
    return x * swell * breath * vol

def bass(freq, dur, vol=1.0):
    n = int(dur * SR)
    t = np.arange(n) / SR
    x = np.sin(2 * np.pi * freq * t) * 0.8
    ph = 2 * np.pi * freq * t
    x += lp(2 * ((ph / (2 * np.pi)) % 1) - 1, 300) * 0.3
    e = np.minimum(1, t / 0.03) * np.exp(-t * 1.4)
    return x * e * vol

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
    out = lp(out, 3600)
    return out * np.exp(-np.arange(len(out)) / SR * 2.6) * vol

# Hijaz D degrees
D2, A2, Bb1, C2 = 73.42, 110.0, 58.27, 65.41
D3, Eb3, Gb3, G3, A3, Bb3, C4, D4 = 146.83, 155.56, 185.0, 196.0, 220.0, 233.08, 261.63, 293.66
B = 60 / 96  # 96 bpm

def samri_bar(t0, intensity=1.0):
    """Gentle samri-feel 4-beat bar: DUM..tik tik DUM.tik | riqq shimmer."""
    place(daff_dum(), t0, 0.85 * intensity)
    place(tik(), t0 + 0.75 * B, 0.7 * intensity)
    place(tik(), t0 + 1.0 * B, 0.55 * intensity)
    place(daff_dum(), t0 + 1.5 * B, 0.7 * intensity)
    place(tik(), t0 + 2.25 * B, 0.7 * intensity)
    place(daff_dum(), t0 + 2.5 * B, 0.5 * intensity)
    place(tik(), t0 + 3.0 * B, 0.6 * intensity)
    place(riqq(), t0 + 3.5 * B, 0.5 * intensity)
    place(soft_clap(), t0 + 2.0 * B, 0.6 * intensity)

# ---------- arrangement ----------
# S1+S2 (0-6.5): tense low pad + lone daff hits on beats
place(pad([D2, A2], 6.8, 0.5, warmth=700), 0)
for at in [0.15, 1.3, 2.65, 3.07, 3.95, 4.37, 5.25, 5.67]:
    place(daff_dum(), at, 0.95)
place(riqq(), 2.0, 0.4); place(riqq(), 4.7, 0.4)

# S3 reveal (6.9-9.5): light burst = warm major-ish lift, oud signature
place(pad([D3, A3, D4], 2.8, 0.55, warmth=1600), 6.95)
place(daff_dum(), 7.55, 1.0)
place(soft_clap(), 7.55, 0.8)
place(bass(D2, 1.4, 0.8), 7.55)
for note, dt, dr in [(D3, 0.0, 0.35), (Gb3, 0.2, 0.35), (A3, 0.4, 0.5), (D4, 0.75, 1.2)]:
    place(pluck(note, dr + 0.3, 0.55), 7.85 + dt)

# S4 hakeem (9.6-19): samri groove + warm pads cycle + light oud
chords = [([D2, A2, D3], D2), ([Bb1, G3*0.5, Bb3*0.25], Bb1), ([C2, G3*0.5, C4*0.25], C2), ([D2, A2, D3], D2)]
tt = 9.6; ci = 0
while tt < 18.8:
    bar_len = 4 * B
    fr, bs = chords[ci % 4]
    place(pad(fr, bar_len + 0.4, 0.4), tt)
    place(bass(bs, 0.8, 0.55), tt)
    place(bass(bs, 0.5, 0.4), tt + 2 * B)
    samri_bar(tt, 0.75)
    tt += bar_len; ci += 1
place(pluck(A3, 0.6, 0.35), 11.9)
place(pluck(D4, 0.9, 0.45), 15.35)
place(pluck(C4, 0.6, 0.3), 15.7)

# S5 services (19-25): groove lifts, brighter pad, hits on cuts
tt = 19.0; ci = 0
while tt < 24.8:
    fr, bs = chords[ci % 4]
    place(pad([f * 2 for f in fr], 4 * B + 0.4, 0.3, warmth=2000), tt)
    place(pad(fr, 4 * B + 0.4, 0.35), tt)
    place(bass(bs, 0.8, 0.65), tt)
    place(bass(bs, 0.5, 0.5), tt + 2 * B)
    samri_bar(tt, 0.95)
    tt += 4 * B; ci += 1
for cut in [19.0, 21.0, 23.0]:
    place(daff_dum(), cut, 1.05)
    place(soft_clap(), cut + 0.05, 0.9)
place(pluck(D4, 0.8, 0.45), 21.3)
place(pluck(A3, 0.8, 0.4), 23.3)

# S6 CTA (25-28.3): elegant climax, groove continues
place(daff_dum(), 25.15, 1.1); place(soft_clap(), 25.15, 0.9); place(bass(D2, 1.2, 0.85), 25.15)
place(pad([D3, A3, D4], 3.6, 0.5, warmth=1800), 25.2)
samri_bar(25.6, 0.85)
samri_bar(25.6 + 4 * B, 0.7)

# WINK OUTRO (28.3-30): groove stops -> ARDAH heritage stamp with Hakeem
place(pluck(D4 * 2, 0.5, 0.5), 28.45)          # wink ting
place(tar_tak(), 28.47, 0.5)
place(ensemble(zeer_hit, 5), 28.95, 1.05)       # ardah double
place(ensemble(lambda: zeer_hit(1.05), 6), 29.25, 1.15)
place(row_clap(), 29.25, 1.0)
place(ensemble(tar_tak, 3, 0.015), 29.5, 0.8)
for note, dt in [(D3, 0), (A3, 0.03), (D4, 0.06)]:
    place(pluck(note, 1.3, 0.7), 29.3 + dt)
place(bass(D2, 1.1, 0.9), 29.25)

# ---------- gentle room ----------
ir_n = int(0.3 * SR)
ir = bp(rng.standard_normal(ir_n) * np.exp(-np.arange(ir_n) / SR * 18), 350, 5500)
track = track + fftconvolve(track, ir)[: len(track)] * 0.045

# ---------- master ----------
track[-int(0.5 * SR):] *= np.linspace(1, 0, int(0.5 * SR))
track = np.tanh(track * 0.9)
track = track / np.max(np.abs(track)) * 0.8
stereo = np.stack([track, track], axis=1)
wavfile.write("../audio/music_v3.wav", SR, (stereo * 32767).astype(np.int16))
print("music_v3.wav written")

"""v7: neutral standard background (Documentary_Tidal) at VERY LOW level,
SFX leading, ardah stamp kept at the wink. No dramatic automation.
"""
import numpy as np
from scipy.signal import butter, sosfilt
from scipy.io import wavfile

SR = 48000
DUR = 30.0
rng = np.random.default_rng(53)

def bp(x, lo, hi):
    return sosfilt(butter(4, [lo, hi], btype="band", fs=SR, output="sos"), x, axis=0)

def lp(x, fc):
    return sosfilt(butter(4, fc, btype="low", fs=SR, output="sos"), x, axis=0)

sr, full = wavfile.read("../audio/downloaded/documentary_tidal.wav")
assert sr == SR
full = full.astype(np.float64) / 32767
cut = full[int(6.0 * SR): int(36.0 * SR)].copy()
n = cut.shape[0]
t = np.arange(n) / SR

vol = np.ones(n)
vol[:int(0.8 * SR)] = np.linspace(0, 1, int(0.8 * SR))          # gentle fade in
m = (t >= 28.3) & (t < 28.7)                                     # small duck for stamp
vol[m] = 1 - (t[m] - 28.3) / 0.4 * 0.6
vol[t >= 28.7] = 0.4
track = cut * vol[:, None]

# ---- ardah stamp (unchanged signature) ----
def env(nn, a, d):
    e = np.ones(nn); na = max(1, int(a * SR))
    e[:na] = np.linspace(0, 1, na)
    return e * np.exp(-np.arange(nn) / SR * d)

def zeer(pitch=1.0):
    nn = int(0.5 * SR); tt = np.arange(nn) / SR
    f = (105 * np.exp(-tt * 16) + 48) * pitch
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR)
    skin = bp(rng.standard_normal(nn), 90, 420) * 0.35
    thud = lp(rng.standard_normal(nn), 160) * 0.5
    return (tone * 1.1 + skin + thud) * env(nn, 0.003, 11)

def tar():
    nn = int(0.12 * SR)
    return bp(rng.standard_normal(nn), 1400, 5200) * env(nn, 0.001, 55) + bp(rng.standard_normal(nn), 500, 1100) * env(nn, 0.001, 45) * 0.5

def ens(fn, layers=5, spread=0.025):
    base = fn(); out = np.zeros(len(base) + int(spread * SR) + 1)
    out[:len(base)] += base
    for _ in range(layers - 1):
        h = fn(); i = int(rng.random() * spread * SR)
        out[i:i + len(h)] += h * (0.55 + 0.35 * rng.random())
    return out / layers * 2.2

def row_clap():
    nn = int(0.3 * SR); x = np.zeros(nn)
    for d in (rng.random(5) * 0.03):
        i = int(d * SR)
        x[i:] += bp(rng.standard_normal(nn - i), 800, 3400) * env(nn - i, 0.001, 34) * (0.6 + 0.4 * rng.random())
    return x * 0.5

def pluck(freq, dur=0.5, v=1.0):
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
    return lp(out, 3600) * np.exp(-np.arange(len(out)) / SR * 2.6) * v

stamp = np.zeros(int(DUR * SR))
def place(s, at, v=1.0):
    i = int(at * SR); j = min(len(stamp), i + len(s))
    stamp[i:j] += s[:j - i] * v

place(pluck(293.66 * 2, 0.5, 0.5), 28.45)
place(tar(), 28.47, 0.5)
place(ens(zeer), 28.95, 1.0)
place(ens(lambda: zeer(1.05), 6), 29.25, 1.1)
place(row_clap(), 29.25, 0.95)
place(ens(tar, 3, 0.015), 29.5, 0.75)
for note, dt in [(146.83, 0), (220.0, 0.03), (293.66, 0.06)]:
    place(pluck(note, 1.3, 0.65), 29.3 + dt)

sr2, sfx = wavfile.read("../audio/sfx_track.wav")
sfx = sfx.astype(np.float64) / 32767

total = int(DUR * SR)
mix = np.zeros((total, 2))
mix[:min(total, len(track))] += track[:total] * 0.11    # music: minimal bed
mix[:min(total, len(sfx))] += sfx[:total] * 0.62
mix += np.stack([stamp, stamp], axis=1) * 0.9
mix[-int(0.4 * SR):] *= np.linspace(1, 0, int(0.4 * SR))[:, None]
mix = np.tanh(mix * 1.0)
mix = mix / np.max(np.abs(mix)) * 0.9
wavfile.write("../audio/final_mix9.wav", SR, (mix * 32767).astype(np.int16))
print("final_mix9.wav written")



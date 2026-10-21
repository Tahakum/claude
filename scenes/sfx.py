"""Synthesize the full 30s SFX track for the Tahakum promo."""
import numpy as np
from scipy.signal import butter, sosfilt
from scipy.io import wavfile

SR = 48000
DUR = 30.0
rng = np.random.default_rng(42)

def env_exp(n, decay):
    t = np.arange(n) / SR
    return np.exp(-t * decay)

def bandpass(x, lo, hi):
    sos = butter(4, [lo, hi], btype="band", fs=SR, output="sos")
    return sosfilt(sos, x)

def lowpass(x, fc):
    sos = butter(4, fc, btype="low", fs=SR, output="sos")
    return sosfilt(sos, x)

def gavel(strength=1.0):
    """Judge hammer: low thump with pitch drop + woody crack."""
    n = int(0.5 * SR)
    t = np.arange(n) / SR
    f = 150 * np.exp(-t * 18) + 55
    thump = np.sin(2 * np.pi * np.cumsum(f) / SR) * env_exp(n, 9)
    crack = bandpass(rng.standard_normal(n), 900, 4000) * env_exp(n, 60) * 0.7
    body = bandpass(rng.standard_normal(n), 150, 600) * env_exp(n, 25) * 0.4
    return (thump * 1.1 + crack + body) * strength

def tick(freq=2100, dur=0.03, vol=1.0):
    n = int(dur * SR)
    t = np.arange(n) / SR
    return np.sin(2 * np.pi * freq * t) * env_exp(n, 160) * vol

def ref_trans(vol=1.0):
    """Transition modeled on the user's reference clip: quick rising sweep
    (~0.25s, 350->5000Hz) -> soft sub tap -> high zing shimmer."""
    n = int(0.55 * SR)
    out = np.zeros(n)
    # rising sweep
    ns = int(0.25 * SR)
    noise = rng.standard_normal(ns)
    seg = max(1, ns // 14)
    sweep = np.zeros(ns)
    for i in range(14):
        a, b = i * seg, min(ns, (i + 1) * seg)
        p = i / 13
        fc = 350 * (5000 / 350) ** p
        sweep[a:b] = bandpass(noise, fc * 0.65, fc * 1.5)[a:b]
    tt = np.linspace(0, 1, ns)
    sweep *= (tt ** 0.7) * np.minimum(1, (1 - tt) * 8 + 0.15)
    out[:ns] += sweep * 0.9
    # soft sub tap at the crest
    nb = int(0.18 * SR)
    tb = np.arange(nb) / SR
    fb = 85 * np.exp(-tb * 18) + 52
    out[ns:ns + nb] += np.sin(2 * np.pi * np.cumsum(fb) / SR) * env_exp(nb, 16) * 0.4
    # high zing shimmer right after the crest
    nz = int(0.16 * SR)
    z = bandpass(rng.standard_normal(nz), 5500, 9000) * env_exp(nz, 28) * 0.3
    tz = np.arange(nz) / SR
    z += np.sin(2 * np.pi * 4200 * tz) * env_exp(nz, 22) * 0.12
    out[ns:ns + nz] += z
    return out * vol

def whoosh(dur=0.7, f0=300, f1=2600, vol=1.0, shape=0.35):
    n = int(dur * SR)
    noise = rng.standard_normal(n)
    out = np.zeros(n)
    seg = max(1, n // 24)
    for i in range(24):
        a, b = i * seg, min(n, (i + 1) * seg)
        p = i / 23
        fc = f0 + (f1 - f0) * p
        out[a:b] = bandpass(noise, max(60, fc * 0.6), fc * 1.6)[a:b]
    t = np.linspace(0, 1, n)
    amp = np.sin(np.pi * t ** shape) ** 2
    return out * amp * vol

def riser(dur=1.4, vol=1.0):
    n = int(dur * SR)
    t = np.arange(n) / SR
    p = t / dur
    f = 90 * (1 + 7 * p ** 2)
    tone = np.sin(2 * np.pi * np.cumsum(f) / SR)
    noi = bandpass(rng.standard_normal(n), 400, 6000) * p ** 2 * 0.5
    amp = p ** 1.6
    return (tone * 0.7 + noi) * amp * vol

def impact(vol=1.0):
    n = int(0.9 * SR)
    t = np.arange(n) / SR
    f = 95 * np.exp(-t * 10) + 42
    sub = np.sin(2 * np.pi * np.cumsum(f) / SR) * env_exp(n, 5)
    punch = bandpass(rng.standard_normal(n), 100, 900) * env_exp(n, 30) * 0.5
    return (sub * 1.2 + punch) * vol

def pop(vol=1.0):
    n = int(0.12 * SR)
    t = np.arange(n) / SR
    f = 500 + 480 * np.linspace(0, 1, n) ** 0.5
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * env_exp(n, 35) * vol

def chime(vol=1.0):
    n = int(0.9 * SR)
    t = np.arange(n) / SR
    a = np.sin(2 * np.pi * 740 * t) * env_exp(n, 6)
    b = np.sin(2 * np.pi * 1108 * t) * env_exp(n, 5)
    b[: int(0.12 * SR)] = 0
    return (a * 0.6 + b * 0.5) * vol

def key_bed(dur=2.2, rate=11, vol=1.0):
    n = int(dur * SR)
    out = np.zeros(n)
    tt = 0.0
    while tt < dur - 0.05:
        kn = int(0.025 * SR)
        k = bandpass(rng.standard_normal(kn), 1200, 5200) * env_exp(kn, 220)
        i = int(tt * SR)
        out[i:i + kn] += k * (0.5 + 0.5 * rng.random())
        tt += (1 / rate) * (0.6 + 0.8 * rng.random())
    return out * vol

def heartbeat(vol=1.0):
    n = int(0.5 * SR)
    t = np.arange(n) / SR
    f = 80 * np.exp(-t * 14) + 45
    b1 = np.sin(2 * np.pi * np.cumsum(f) / SR) * env_exp(n, 12)
    out = np.zeros(int(0.8 * SR))
    out[:n] += b1
    out[int(0.22 * SR):int(0.22 * SR) + n] += b1 * 0.7
    return out * vol

track = np.zeros(int(DUR * SR))

def place(sound, at, vol=1.0):
    i = int(at * SR)
    j = min(len(track), i + len(sound))
    track[i:j] += sound[: j - i] * vol

# ---- S1 hook ----
place(gavel(1.0), 0.10)
place(ref_trans(0.45), 0.0)
place(pop(0.8), 1.25)
# ---- S2 pain beats ----
tt = 2.70
step = 0.52
while tt < 6.2:
    place(tick(1900 if int(tt * 2) % 2 == 0 else 1500, vol=0.4), tt)
    step = max(0.3, step * 0.93)
    tt += step
for w1, w2 in [(2.65, 3.07), (3.95, 4.37), (5.25, 5.67)]:
    place(ref_trans(0.4), w1 - 0.05)
    place(impact(0.5), w2)
# ---- S3 reveal (softer than ordinary transitions) ----
place(riser(1.3, 0.2), 6.6)
place(impact(0.2), 7.5)
place(gavel(0.15), 7.52)
place(whoosh(0.5, 800, 4000, 0.28), 7.9)
# ---- S4 hakeem ----
place(ref_trans(0.55), 9.5)
place(pop(0.7), 10.6)
place(pop(0.5), 11.55)
place(key_bed(2.1, 12, 0.8), 11.8)
place(pop(0.6), 14.15)
place(chime(0.7), 15.25)
place(pop(0.6), 15.7)
# ---- S5 services ----
for at in [19.0, 21.0, 23.0]:
    place(ref_trans(0.5), at - 0.1)
    place(impact(0.35), at + 0.25)
# ---- S6 CTA ----
place(riser(0.9, 0.12), 24.5)
place(impact(0.2), 25.15)
place(gavel(0.12), 25.5)
place(pop(0.6), 26.05)
place(ref_trans(0.5), 26.2)
place(pop(0.45), 28.42)
place(chime(0.4), 29.3)

# master: gentle fade-out, normalize
track[-int(0.5 * SR):] *= np.linspace(1, 0, int(0.5 * SR))
track = np.tanh(track * 1.1)
track = track / np.max(np.abs(track)) * 0.85
stereo = np.stack([track, track], axis=1)
wavfile.write("../audio/sfx_track.wav", SR, (stereo * 32767).astype(np.int16))
print("sfx_track.wav written", stereo.shape)

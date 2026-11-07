"""Charts for the README. python make_charts.py"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import model
from model import FORECAST

plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10,
                     "axes.spines.top": False, "axes.spines.right": False})
INK, MUTED, BLUE, RED = "#222222", "#888888", "#2f5d8a", "#b5452f"


def value_bridge(res, path="value_bridge.png"):
    s = res["summary"]
    steps = [("DCF of existing\nbusiness", s["dcf_ev"]), ("JVs at book", s["jv_book"]),
             ("Synergies\n(PV, after tax)", s["synergy_pv"]), ("Pipeline /\nunexplained", s["gap"])]
    total = s["deal_ev"] + s["fees"]

    fig, ax = plt.subplots(figsize=(8, 4.2))
    bottom = 0
    for i, (label, v) in enumerate(steps):
        colour = RED if i == len(steps) - 1 else BLUE
        ax.bar(i, v, bottom=bottom, color=colour, width=0.6)
        ax.text(i, bottom + v / 2, f"{v:.0f}", ha="center", va="center", color="white", fontweight="bold")
        bottom += v
    ax.bar(len(steps), total, color=INK, width=0.6)
    ax.text(len(steps), total / 2, f"{total:.0f}", ha="center", va="center", color="white", fontweight="bold")
    ax.set_xticks(range(len(steps) + 1), [l for l, _ in steps] + ["Price paid\n(EV + fees)"])
    ax.set_ylabel("EUR m")
    ax.set_title("What a 25% premium pays for", loc="left", color=INK)
    fig.tight_layout()
    fig.savefig(path, dpi=160)


def accretion_vs_price(path="accretion_vs_price.png"):
    prices = list(range(35, 81, 5))
    base = model.Assumptions().price_mwh

    def flat(p):
        return {y: (base[y] if y < 2027 else p) for y in FORECAST}

    fig, ax = plt.subplots(figsize=(8, 3.8))
    for cut, colour in [(0.0, MUTED), (0.6, BLUE), (0.9, INK)]:
        acc = [model.accretion(2028, price_mwh=flat(p), gna_cut=cut) * 100 for p in prices]
        ax.plot(prices, acc, color=colour, marker="o", ms=3, label=f"{cut:.0%} of Orrön G&A removed")
    ax.axhline(0, color=INK, lw=0.8)
    ax.axvline(base[2028], color=MUTED, lw=0.8, ls="--")
    ax.text(base[2028] + 0.5, ax.get_ylim()[0] * 0.9, "base case", color=MUTED, fontsize=9)
    ax.set_xlabel("Achieved power price from 2027, EUR/MWh")
    ax.set_ylabel("2028E EPS accretion, %")
    ax.set_title("Cloudberry EPS impact depends almost entirely on the power price", loc="left", color=INK)
    ax.legend(frameon=False, fontsize=9)
    fig.tight_layout()
    fig.savefig(path, dpi=160)


if __name__ == "__main__":
    res = model.run()
    value_bridge(res)
    accretion_vs_price()

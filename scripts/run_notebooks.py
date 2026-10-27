"""Run the notebooks headless, in parallel, and save the executed copies to results/.

Usage:
    python3 scripts/run_notebooks.py          # all four
    python3 scripts/run_notebooks.py 1 3      # just Q1 and Q3

Each notebook runs in its own Jupyter kernel with results/ as its working directory, so the
downloaded cross sections (kaeri_*.txt) are cached there and reused on the next run. The executed
notebook (with all plots and printed numbers) is written to results/<name>.ipynb and each plot to
results/figures/Q<n>_plot<k>.png. On a laptop, keep the machine awake for long runs
(macOS: `caffeinate -i python3 scripts/run_notebooks.py`).
"""
import base64
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor

import nbformat
from nbclient import NotebookClient

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "results")
NOTEBOOKS = {1: "Q1_pure_U238.ipynb", 2: "Q2_natural_uranium.ipynb",
             3: "Q3_enriched_uranium_light_water.ipynb", 4: "Q4_natural_uranium_heavy_water.ipynb"}


def run_one(q):
    nb = nbformat.read(os.path.join(ROOT, "notebooks", NOTEBOOKS[q]), as_version=4)
    t = time.time()
    try:
        NotebookClient(nb, timeout=-1, kernel_name="python3", resources={"metadata": {"path": OUT}}).execute()
        ok = True
    except Exception as e:
        ok = False
        print(f"Q{q} FAILED: {e}", flush=True)
    nbformat.write(nb, os.path.join(OUT, NOTEBOOKS[q]))
    k = 0
    for c in nb.cells:
        for o in c.get("outputs", []) if c.cell_type == "code" else []:
            if "data" in o and "image/png" in o["data"]:
                k += 1
                with open(os.path.join(OUT, "figures", f"Q{q}_plot{k}.png"), "wb") as f:
                    f.write(base64.b64decode(o["data"]["image/png"]))
    return f"Q{q} {'OK' if ok else 'FAILED'} in {(time.time() - t) / 60:.1f} min, {k} figures"


if __name__ == "__main__":
    qs = [int(a) for a in sys.argv[1:]] or list(NOTEBOOKS)
    os.makedirs(os.path.join(OUT, "figures"), exist_ok=True)
    with ProcessPoolExecutor(len(qs)) as pool:
        for line in pool.map(run_one, qs):
            print(line, flush=True)

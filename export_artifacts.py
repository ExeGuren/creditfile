"""
export_artifacts.py
-------------------
Exports the LightGBM model and QuantileTransformer scaler to JSON files
that can be consumed by the static JavaScript pipeline.

Run once:
    python export_artifacts.py

Outputs:
    docs/artifacts/model.json        — LightGBM tree ensemble
    docs/artifacts/scaler.json       — QuantileTransformer lookup tables
    docs/artifacts/bow.json          — Bag-of-words vocabulary
"""

import json
import os
import joblib
import lightgbm as lgb
import numpy as np

OUT_DIR = os.path.join(os.path.dirname(__file__), 'docs', 'artifacts')
os.makedirs(OUT_DIR, exist_ok=True)

ARTIFACTS = os.path.join(os.path.dirname(__file__), 'creditfile', 'artifacts')

# ── 1. LightGBM model ─────────────────────────────────────────────────────────
print("Exporting LightGBM model...")
booster = lgb.Booster(model_file=os.path.join(ARTIFACTS, 'model.txt'))
model_dump = booster.dump_model()

# We only need the tree structures for raw score inference
trees = model_dump['tree_info']

def export_node(node):
    """Recursively convert a LightGBM node dict to a compact structure."""
    if 'leaf_value' in node:
        return {'leaf': node['leaf_value']}
    return {
        'feature':    node['split_feature'],
        'threshold':  node['threshold'],
        'default_left': node.get('default_left', False),
        'left':       export_node(node['left_child']),
        'right':      export_node(node['right_child']),
    }

exported_trees = []
for t in trees:
    exported_trees.append({
        'shrinkage': t.get('shrinkage', 1.0),
        'tree':      export_node(t['tree_structure'])
    })

model_json = {
    'average_output': model_dump.get('average_output', False),
    'trees': exported_trees
}

with open(os.path.join(OUT_DIR, 'model.json'), 'w') as f:
    json.dump(model_json, f, separators=(',', ':'))
print(f"  → {len(exported_trees)} trees exported")

# ── 2. QuantileTransformer scaler ─────────────────────────────────────────────
print("Exporting QuantileTransformer scaler...")
scaler = joblib.load(os.path.join(ARTIFACTS, 'score-scaler.pickle'))

scaler_json = {
    'quantiles':  scaler.quantiles_[:, 0].tolist(),
    'references': scaler.references_.tolist(),
    'output_distribution': scaler.output_distribution,
}

with open(os.path.join(OUT_DIR, 'scaler.json'), 'w') as f:
    json.dump(scaler_json, f, separators=(',', ':'))
print(f"  → {len(scaler_json['quantiles'])} quantile points exported")

# ── 3. Bag-of-words vocabulary ────────────────────────────────────────────────
print("Exporting BoW vocabulary...")
bow = joblib.load(os.path.join(ARTIFACTS, 'bow.pickle'))
vocabulary = {word: int(idx) for word, idx in bow.vocabulary_.items()}
feature_names = list(bow.get_feature_names_out())

bow_json = {
    'vocabulary':    vocabulary,
    'feature_names': feature_names,
}

with open(os.path.join(OUT_DIR, 'bow.json'), 'w') as f:
    json.dump(bow_json, f, separators=(',', ':'))
print(f"  → {len(vocabulary)} BoW tokens exported")

print("\nDone. Files written to:", OUT_DIR)

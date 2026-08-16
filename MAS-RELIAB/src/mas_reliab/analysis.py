"""Measured summaries, paired inference, hypothesis decisions, figures, and tables."""
from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from scipy import stats

from .utils import stable_seed

sns.set_theme(style="whitegrid", context="paper")
TOPOLOGY_ORDER = ["single", "sequential", "parallel", "hierarchical"]


def _cluster_ci(frame: pd.DataFrame, value: str, cluster: str = "task_id", reps: int = 2000) -> tuple[float, float, float]:
    clean = frame[[cluster, value]].dropna()
    if clean.empty:
        return math.nan, math.nan, math.nan
    by_cluster = clean.groupby(cluster)[value].mean()
    estimate = float(clean[value].mean())
    if len(by_cluster) == 1:
        return estimate, estimate, estimate
    local = np.random.default_rng(stable_seed("bootstrap", value, len(clean), reps))
    values = by_cluster.to_numpy(float)
    means = np.empty(reps)
    for i in range(reps):
        means[i] = local.choice(values, size=len(values), replace=True).mean()
    low, high = np.quantile(means, [0.025, 0.975])
    return estimate, float(low), float(high)


def _summary(frame: pd.DataFrame, groups: list[str], metrics: list[str], reps: int) -> pd.DataFrame:
    rows = []
    grouper: Any = groups[0] if len(groups) == 1 else groups
    for keys, subset in frame.groupby(grouper, dropna=False):
        keys = (keys,) if len(groups) == 1 else tuple(keys)
        row = dict(zip(groups, keys))
        row["n_episodes"] = len(subset)
        row["n_tasks"] = subset.task_id.nunique() if "task_id" in subset else math.nan
        for metric in metrics:
            estimate, low, high = _cluster_ci(subset, metric, reps=reps)
            row[metric] = estimate
            row[f"{metric}_ci_low"] = low
            row[f"{metric}_ci_high"] = high
        rows.append(row)
    return pd.DataFrame(rows)


def _paired(frame: pd.DataFrame, key: list[str], treatment_col: str, a: str, b: str, outcome: str) -> dict[str, Any]:
    sub = frame[frame[treatment_col].isin([a, b])]
    wide = sub.pivot_table(index=key, columns=treatment_col, values=outcome, aggfunc="mean").dropna()
    if wide.empty or a not in wide or b not in wide:
        return {"contrast": f"{a} - {b}", "outcome": outcome, "n_pairs": 0, "mean_difference": math.nan,
                "ci_low": math.nan, "ci_high": math.nan, "p_value": math.nan, "effect_size": math.nan,
                "test": "unavailable"}
    diff = (wide[a] - wide[b]).to_numpy(float)
    local = np.random.default_rng(stable_seed("paired", a, b, outcome, len(diff)))
    boot = np.array([local.choice(diff, size=len(diff), replace=True).mean() for _ in range(2000)])
    low, high = np.quantile(boot, [0.025, 0.975])
    nonzero = diff[diff != 0]
    if set(np.unique(diff)).issubset({-1.0, 0.0, 1.0}):
        positive = int((diff > 0).sum())
        negative = int((diff < 0).sum())
        discordant = positive + negative
        p = float(stats.binomtest(min(positive, negative), discordant, 0.5).pvalue) if discordant else 1.0
        effect = (positive - negative) / discordant if discordant else 0.0
        test = "exact McNemar/binomial sign test"
    elif len(nonzero) >= 2:
        try:
            p = float(stats.wilcoxon(diff, alternative="two-sided", zero_method="wilcox").pvalue)
        except ValueError:
            p = 1.0
        ranks = stats.rankdata(abs(nonzero))
        effect = float((ranks[nonzero > 0].sum() - ranks[nonzero < 0].sum()) / ranks.sum())
        test = "paired Wilcoxon signed-rank"
    else:
        p, effect, test = 1.0, 0.0, "paired sign (insufficient nonzero pairs)"
    return {"contrast": f"{a} - {b}", "outcome": outcome, "n_pairs": len(diff),
            "mean_difference": float(diff.mean()), "ci_low": float(low), "ci_high": float(high),
            "p_value": p, "effect_size": effect, "test": test}


def _holm(p_values: list[float]) -> list[float]:
    result = [math.nan] * len(p_values)
    valid = [(i, p) for i, p in enumerate(p_values) if not math.isnan(p)]
    ordered = sorted(valid, key=lambda pair: pair[1])
    running = 0.0
    m = len(ordered)
    for rank, (index, p) in enumerate(ordered):
        adjusted = min(1.0, (m - rank) * p)
        running = max(running, adjusted)
        result[index] = running
    return result


def _dataset_table(tasks_df: pd.DataFrame) -> pd.DataFrame:
    table = tasks_df.groupby(["split", "family"]).size().rename("n_tasks").reset_index()
    return table


def _baseline_table(e1: pd.DataFrame, reps: int) -> pd.DataFrame:
    summary = _summary(e1, ["topology"], ["task_success", "work_units", "latency_ms"], reps)
    pass_rows = []
    for topology, subset in e1.groupby("topology"):
        grouped = subset.groupby("task_id")
        pass_k = grouped.task_success.all().mean()
        consistency = grouped.final_output.apply(lambda x: x.astype(str).nunique() == 1).mean()
        pass_rows.append({"topology": topology, "pass_k": pass_k, "exact_output_consistency": consistency})
    return summary.merge(pd.DataFrame(pass_rows), on="topology")


def _tradeoff_table(episodes: pd.DataFrame, reps: int) -> pd.DataFrame:
    intervention = episodes[episodes.experiment.isin(["E4_verification_ablation", "E5_recovery"])].copy()
    return _summary(intervention, ["experiment", "condition", "topology"],
                    ["task_success", "epr", "daf", "work_units", "latency_ms", "recovered"], reps)


def _save_figures(tables: dict[str, pd.DataFrame], figure_dir: Path) -> None:
    figure_dir.mkdir(parents=True, exist_ok=True)
    palette = dict(zip(TOPOLOGY_ORDER, sns.color_palette("colorblind", 4)))

    base = tables["baseline"]
    fig, axes = plt.subplots(1, 3, figsize=(11, 3.4))
    for ax, metric, title in zip(axes, ["task_success", "pass_k", "exact_output_consistency"],
                                 ["Task success", "Pass-all repeats", "Exact-output consistency"]):
        sns.barplot(data=base, x="topology", y=metric, hue="topology", order=TOPOLOGY_ORDER,
                    hue_order=TOPOLOGY_ORDER, palette=palette, legend=False, ax=ax)
        ax.set_ylim(0, 1.04); ax.set_title(title); ax.tick_params(axis="x", rotation=25); ax.set_xlabel("")
    fig.suptitle("E1 — Baseline reliability dimensions (simulation pilot)")
    fig.tight_layout(); fig.savefig(figure_dir / "figure_01_baseline_reliability.png", dpi=220); plt.close(fig)

    prop = tables["propagation"]
    fig, axes = plt.subplots(1, 2, figsize=(9, 3.7))
    sns.pointplot(data=prop, x="fault_position", y="epr", hue="topology", hue_order=TOPOLOGY_ORDER,
                  order=["early", "middle", "late"], palette=palette, errorbar=None, ax=axes[0])
    sns.pointplot(data=prop, x="fault_position", y="daf", hue="topology", hue_order=TOPOLOGY_ORDER,
                  order=["early", "middle", "late"], palette=palette, errorbar=None, ax=axes[1])
    axes[0].set_title("Event propagation rate"); axes[1].set_title("Downstream affected final")
    axes[1].get_legend().remove(); axes[0].legend(title="Topology", fontsize=8)
    fig.suptitle("E2 — Propagation by functional injection position")
    fig.tight_layout(); fig.savefig(figure_dir / "figure_02_fault_propagation.png", dpi=220); plt.close(fig)

    attr = tables["attribution"]
    fig, ax = plt.subplots(figsize=(6.5, 3.7))
    sns.barplot(data=attr, x="observability", y="top1_correct",
                order=["output_only", "sparse_trace", "full_trace"], color="#3274A1", ax=ax)
    ax.set_ylim(0, 1.04); ax.set_ylabel("Top-1 attribution accuracy"); ax.set_xlabel("Evidence view")
    ax.set_title("E3 — Attribution improves with trace evidence")
    fig.tight_layout(); fig.savefig(figure_dir / "figure_03_attribution.png", dpi=220); plt.close(fig)

    ver = tables["verification"]
    fig, axes = plt.subplots(1, 2, figsize=(10, 3.7))
    sns.barplot(data=ver, x="condition", y="task_success", hue="topology", hue_order=TOPOLOGY_ORDER,
                palette=palette, ax=axes[0])
    sns.barplot(data=ver, x="condition", y="epr", hue="topology", hue_order=TOPOLOGY_ORDER,
                palette=palette, ax=axes[1])
    for ax in axes: ax.tick_params(axis="x", rotation=30); ax.set_xlabel("")
    axes[0].set_title("Task success"); axes[1].set_title("Propagation")
    axes[1].get_legend().remove(); axes[0].legend(fontsize=7)
    fig.suptitle("E4 — Verification ablation")
    fig.tight_layout(); fig.savefig(figure_dir / "figure_04_verification.png", dpi=220); plt.close(fig)

    rec = tables["recovery"]
    fig, axes = plt.subplots(1, 2, figsize=(10, 3.7))
    sns.barplot(data=rec, x="condition", y="recovered", hue="topology", hue_order=TOPOLOGY_ORDER,
                palette=palette, ax=axes[0])
    sns.barplot(data=rec, x="condition", y="work_units", hue="topology", hue_order=TOPOLOGY_ORDER,
                palette=palette, ax=axes[1])
    for ax in axes: ax.tick_params(axis="x", rotation=30); ax.set_xlabel("")
    axes[0].set_title("Recovery rate"); axes[1].set_title("Simulated work units")
    axes[1].get_legend().remove(); axes[0].legend(fontsize=7)
    fig.suptitle("E5 — Reliability–cost recovery trade-off")
    fig.tight_layout(); fig.savefig(figure_dir / "figure_05_recovery.png", dpi=220); plt.close(fig)

    trade = tables["tradeoff"]
    fig, ax = plt.subplots(figsize=(7.2, 4.5))
    sns.scatterplot(data=trade, x="work_units", y="task_success", hue="topology", style="condition",
                    hue_order=TOPOLOGY_ORDER, palette=palette, s=80, ax=ax)
    ax.set_title("E6 — Measured reliability–cost operating points")
    ax.set_xlabel("Mean simulated work units (lower is better)"); ax.set_ylabel("Task success (higher is better)")
    ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", fontsize=6)
    fig.tight_layout(); fig.savefig(figure_dir / "figure_06_tradeoff.png", dpi=220); plt.close(fig)


def analyze(episodes: pd.DataFrame, attribution: pd.DataFrame, tasks_df: pd.DataFrame,
            config: dict[str, Any], root: Path) -> dict[str, pd.DataFrame]:
    reps = int(config["analysis"]["bootstrap_replicates"])
    result_root = root / config["outputs"]["root"]
    table_dir, analysis_dir, figure_dir = result_root / "tables", result_root / "analysis", result_root / "figures"
    table_dir.mkdir(parents=True, exist_ok=True); analysis_dir.mkdir(parents=True, exist_ok=True)

    e1 = episodes[episodes.experiment == "E1_baseline_reliability"].copy()
    e2_all = episodes[episodes.experiment == "E2_fault_propagation"].copy()
    e2 = e2_all[e2_all.injection_consumed].copy()
    e4 = episodes[(episodes.experiment == "E4_verification_ablation") & episodes.injection_consumed].copy()
    e5 = episodes[(episodes.experiment == "E5_recovery") & episodes.injection_consumed].copy()

    tables: dict[str, pd.DataFrame] = {}
    tables["dataset"] = _dataset_table(tasks_df)
    tables["baseline"] = _baseline_table(e1, reps)
    tables["injection_integrity"] = _summary(e2_all, ["topology", "fault_type", "fault_severity"],
        ["injection_applied", "injection_consumed"], reps)
    tables["propagation"] = _summary(e2, ["topology", "fault_type", "fault_severity", "fault_position"],
        ["task_success", "epr", "daf", "amplification", "propagation_depth", "work_units"], reps)
    tables["attribution"] = _summary(attribution,
        ["observability"], ["top1_correct", "reciprocal_rank"], reps)
    tables["verification"] = _summary(e4, ["topology", "condition"],
        ["task_success", "epr", "daf", "detected", "recovered", "work_units", "latency_ms"], reps)
    tables["recovery"] = _summary(e5, ["topology", "condition"],
        ["task_success", "epr", "daf", "detected", "recovered", "work_units", "latency_ms"], reps)
    tables["tradeoff"] = _tradeoff_table(episodes, reps)

    filenames = {
        "dataset": "table_01_dataset.csv", "baseline": "table_02_baseline.csv",
        "injection_integrity": "table_03_injection_integrity.csv",
        "propagation": "table_04_propagation.csv", "attribution": "table_05_attribution.csv",
        "verification": "table_06_verification.csv", "recovery": "table_07_recovery.csv",
        "tradeoff": "table_08_tradeoff.csv",
    }
    for name, table in tables.items():
        table.to_csv(table_dir / filenames[name], index=False)

    tests: list[dict[str, Any]] = []
    # H1: consumed injections, functional early vs late within task/fault/repeat/topology.
    h1_frame = e2.copy()
    h1_frame["pair"] = h1_frame.block_id.str.replace(r"::(early|middle|late)::", "::POSITION::", regex=True)
    for outcome in ["daf", "propagation_depth", "epr"]:
        row = _paired(h1_frame, ["task_id", "topology", "fault_type", "fault_severity"], "fault_position", "early", "late", outcome)
        row.update({"hypothesis": "H1", "family": "position"}); tests.append(row)
    for outcome in ["daf", "epr", "task_success"]:
        row = _paired(e2, ["task_id", "topology", "fault_type", "fault_position"],
                      "fault_severity", 2, 1, outcome)
        row.update({"hypothesis": "RQ2-exploratory", "family": "severity"}); tests.append(row)
    # H2: same fault episode under different evidence views.
    for comparison in [("full_trace", "output_only"), ("full_trace", "sparse_trace")]:
        row = _paired(attribution, ["episode_id"], "observability", comparison[0], comparison[1], "top1_correct")
        row.update({"hypothesis": "H2", "family": "observability"}); tests.append(row)
    # H3: local checking against final-only and none.
    for outcome in ["epr", "daf", "task_success", "work_units"]:
        for b in ["verify=final;recovery=retry", "verify=none;recovery=none"]:
            row = _paired(e4, ["task_id", "topology", "fault_type"], "condition",
                          "verify=local;recovery=retry", b, outcome)
            row.update({"hypothesis": "H3", "family": "verification"}); tests.append(row)
    # H4: alternate/redundant and isolate against retry.
    for strategy in ["recovery=isolate", "recovery=redundant"]:
        for outcome in ["recovered", "task_success", "work_units"]:
            row = _paired(e5, ["task_id", "topology"], "condition", strategy, "recovery=retry", outcome)
            row.update({"hypothesis": "H4", "family": "recovery"}); tests.append(row)
    # H5: hierarchical manager versus leaf; parallel versus hierarchical isolated worker.
    hierarchy = e2[e2.topology == "hierarchical"]
    for outcome in ["daf", "epr"]:
        row = _paired(hierarchy, ["task_id", "fault_type", "fault_severity"], "fault_position", "middle", "early", outcome)
        row.update({"hypothesis": "H5", "family": "topology_location"}); tests.append(row)
    early = e2[e2.fault_position == "early"]
    row = _paired(early, ["task_id", "fault_type", "fault_severity"], "topology", "parallel", "hierarchical", "daf")
    row.update({"hypothesis": "H5", "family": "topology_location"}); tests.append(row)
    # H6: compare topology rankings and directly test the TSR leader against the containment leader.
    success_rank = tables["baseline"].sort_values("task_success", ascending=False).topology.tolist()
    cost_rank = tables["baseline"].sort_values("work_units", ascending=True).topology.tolist()
    e2_rank = e2.groupby("topology").daf.mean().sort_values().index.tolist()
    h6_containment = _paired(e2, ["task_id", "fault_type", "fault_severity", "fault_position"], "topology",
                             success_rank[0], e2_rank[0], "daf")
    h6_containment.update({"hypothesis": "H6", "family": "rank_discordance"}); tests.append(h6_containment)
    if success_rank[0] != cost_rank[0]:
        h6_cost = _paired(e1, ["task_id"], "topology", success_rank[0], cost_rank[0], "work_units")
        h6_cost.update({"hypothesis": "H6", "family": "rank_discordance"}); tests.append(h6_cost)

    tests_df = pd.DataFrame(tests)
    tests_df["p_adjusted_holm"] = _holm(tests_df.p_value.tolist())
    tests_df.to_csv(analysis_dir / "statistical_tests.csv", index=False)

    alpha = float(config["analysis"]["alpha"])
    def sig(h: str, contains: str, direction: str, outcome: str | None = None) -> bool:
        subset = tests_df[(tests_df.hypothesis == h) & tests_df.contrast.str.contains(contains, regex=False)]
        if outcome is not None:
            subset = subset[subset.outcome == outcome]
        if subset.empty: return False
        directional = subset.mean_difference.gt(0) if direction == "positive" else subset.mean_difference.lt(0)
        return bool((directional & subset.p_adjusted_holm.lt(alpha)).any())

    h1_depth = sig("H1", "early - late", "positive", "propagation_depth")
    h1_epr = sig("H1", "early - late", "positive", "epr")
    h1_daf_contradiction = sig("H1", "early - late", "negative", "daf")
    h1_decision = "supported" if (h1_depth and h1_epr and not h1_daf_contradiction) else ("mixed support" if (h1_depth or h1_epr) else "not supported")
    h2_support = sig("H2", "full_trace - output_only", "positive", "top1_correct")
    h3_support = sig("H3", "local;recovery=retry - verify=final", "negative", "epr") or sig("H3", "local;recovery=retry - verify=final", "negative", "daf")
    h4_recovery = sig("H4", "recovery=redundant - recovery=retry", "positive", "recovered") or sig("H4", "recovery=isolate - recovery=retry", "positive", "recovered")
    h4_cost = sig("H4", "recovery=redundant - recovery=retry", "positive", "work_units") or sig("H4", "recovery=isolate - recovery=retry", "positive", "work_units")
    h5_manager = sig("H5", "middle - early", "positive")
    h5_parallel = sig("H5", "parallel - hierarchical", "negative")

    rank_discordance = success_rank != cost_rank or success_rank != e2_rank
    h6_support = rank_discordance and sig("H6", f"{success_rank[0]} - {e2_rank[0]}", "positive", "daf")

    decisions = pd.DataFrame([
        {"hypothesis": "H1", "decision": h1_decision,
         "observed_basis": "Early injections increased EPR and depth, but their DAF contrast was lower than late injections; the directional evidence was mixed."},
        {"hypothesis": "H2", "decision": "supported" if h2_support else "not supported",
         "observed_basis": "Paired full-trace versus output-only top-1 attribution."},
        {"hypothesis": "H3", "decision": "supported" if h3_support else "not supported",
         "observed_basis": "Paired local-verification reduction in EPR/DAF versus final-only/no verification."},
        {"hypothesis": "H4", "decision": "supported" if (h4_recovery and h4_cost) else "not supported",
         "observed_basis": "Recovery gain and work-unit cost of isolate/redundant versus retry."},
        {"hypothesis": "H5", "decision": "supported" if (h5_manager and h5_parallel) else "not supported",
         "observed_basis": "Hierarchy manager-versus-leaf and parallel-versus-hierarchical worker contrasts."},
        {"hypothesis": "H6", "decision": "supported" if h6_support else "not supported",
         "observed_basis": f"TSR rank={success_rank}; cost rank={cost_rank}; DAF-containment rank={e2_rank}."},
    ])
    decisions.to_csv(analysis_dir / "hypothesis_decisions.csv", index=False)

    protocol = {
        "scheduled_fault_episodes": int(e2_all.injection_scheduled.sum()),
        "applied_fault_episodes": int(e2_all.injection_applied.sum()),
        "consumed_fault_episodes": int(e2_all.injection_consumed.sum()),
        "application_rate": float(e2_all.injection_applied.mean()),
        "consumption_rate_given_scheduled": float(e2_all.injection_consumed.mean()),
        "analysis_denominator": "Consumed injections for propagation, verification, and recovery effect summaries",
    }
    with open(analysis_dir / "analysis_manifest.json", "w", encoding="utf-8") as handle:
        json.dump(protocol, handle, indent=2)

    lines = ["# RQ1–RQ6 measured answers", "", "All answers are limited to the configured offline simulator pilot.", ""]
    for _, row in decisions.iterrows():
        rq = {"H6":"RQ1", "H1":"RQ2", "H2":"RQ3", "H3":"RQ4", "H4":"RQ5", "H5":"RQ6"}[row.hypothesis]
        lines += [f"## {rq} / {row.hypothesis}: {row.decision}", row.observed_basis, ""]
    (analysis_dir / "rq_answers.md").write_text("\n".join(lines), encoding="utf-8")

    _save_figures(tables, figure_dir)
    return tables

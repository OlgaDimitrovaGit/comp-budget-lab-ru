# -*- coding: utf-8 -*-
"""Reference numbers for page/check-engine.mjs: reference_calc.analyse() on each scenario, as JSON.

    python dump_reference.py scenarios.json > reference.json

A scenario is {"data", "bands", "settings"}; settings use the page's keys
(effective {year, month}, ratings {"1": [target, min]}).
"""
import json
import os
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from reference_calc import analyse, load, load_bands  # noqa: E402


def settings(s):
    out = dict(s)
    if "effective" in out:
        e = out.pop("effective")
        out["effective_date"] = date(e["year"], e["month"], 1)
    if "ratings" in out:
        out["ratings"] = {int(k): tuple(v) for k, v in out["ratings"].items()}
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    scenarios = json.load(open(sys.argv[1], encoding="utf-8"))
    out = []
    for sc in scenarios:
        r = analyse(load(os.path.join(HERE, sc["data"])), load_bands(os.path.join(HERE, sc["bands"])),
                    settings(sc.get("settings", {})))
        keys = ("pay_before", "contrib_before", "payroll", "eligible_base", "merit_pct", "dev_pp", "dev_rub",
                "dev_rub_in_year", "cost_annual", "cost_in_year", "load_full_pct", "load_in_year_pct",
                "below_min_after", "red_circled")
        out.append({
            "name": sc["name"],
            **{k: r[k] for k in keys},
            **{t: r[t] for t in ("merit", "target")},
            "emp": {e["id"]: {k: (int(e[k]) if isinstance(e[k], bool) else e[k])
                              for k in ("increase", "new_base", "eligible", "zone")}
                    for e in r["emp"]},
        })
    json.dump(out, sys.stdout, ensure_ascii=False)


if __name__ == "__main__":
    main()

"""Generate a realistic, deliberately messy pharmacy supplier-order dataset.

The data mimics what an operations team receives from multiple suppliers:
mixed date formats, currency strings, inconsistent casing, stray whitespace,
missing cells, and exact duplicate rows. It is generated deterministically so
demos and tests are repeatable.
"""

from __future__ import annotations

import random
from datetime import date, timedelta

import pandas as pd

SUPPLIERS = [
    "Müller Pharma GmbH",
    "Café Médical SARL",
    "Northwind Medical Supply",
    "Sanité Distribution",
    "Evergreen Health Ltd",
    "Bergström & Co.",
    "PureCare Wholesale",
    "Apex Biopharma",
    "Lakeside Pharma",
    "Vitalis Import Export",
]

PRODUCTS = [
    ("Amoxicillin 500mg", "Prescription", 12.50),
    ("Ibuprofen 200mg", "OTC", 4.20),
    ("Vitamin D3 1000 IU", "OTC", 8.75),
    ("Insulin Pen 3ml", "Cold Chain", 42.00),
    ("Surgical Gloves (Box)", "Supplies", 18.40),
    ("Digital Thermometer", "Equipment", 26.90),
    ("Saline IV 500ml", "Prescription", 6.30),
    ("Face Masks (50 pack)", "Supplies", 14.00),
    ("Allergy Relief 10mg", "OTC", 9.60),
    ("Blood Pressure Monitor", "Equipment", 64.50),
    ("Hydration Sachets", "OTC", 5.25),
    ("Medical Tape 5m", "Supplies", 3.80),
]

STATUSES = ["Delivered", "Delivered", "Delivered", "Processing", "Shipped", "Backordered", "Cancelled"]
REGIONS = ["Northeast", "Southeast", "Midwest", "West", "International"]
DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y", "%b %d, %Y")

COLUMNS = [
    "  order id  ",
    "order_date",
    "supplier name",
    "product name",
    "category",
    "qty",
    "unit price ($)",
    "total amount",
    "delivery_status",
    "warehouse region",
]


def _case_variant(text: str, rng: random.Random) -> str:
    """Return a casing/whitespace variant of ``text``."""
    roll = rng.random()
    if roll < 0.12:
        return f"  {text.upper()}  "
    if roll < 0.22:
        return f" {text.lower()} "
    if roll < 0.32:
        return f" {text} "
    return text


def _maybe_missing(value, rng: random.Random, probability: float = 0.04):
    roll = rng.random()
    if roll < probability:
        return None
    if roll < probability + 0.015:
        return ""
    return value


def _currency(value: float, rng: random.Random) -> str:
    """Render a number using one of several plausible messy formats."""
    style = rng.randrange(5)
    if style == 0:
        return f"${value:,.2f}"
    if style == 1:
        return f" ${value:,.2f} "
    if style == 2:
        return f"USD {value:,.2f}"
    if style == 3:
        return f"${value:,.2f}  "
    return f"{value:,.2f}"


def generate_sample_data(rows: int = 120, seed: int = 7) -> pd.DataFrame:
    """Return a messy supplier-order DataFrame.

    Args:
        rows: Number of unique orders before duplicates are injected.
        seed: Random seed for reproducible sample data.
    """
    if rows < 10:
        raise ValueError("rows must be at least 10")

    rng = random.Random(seed)
    start = date(2026, 1, 5)
    records = []

    for index in range(rows):
        product, category, base_price = rng.choice(PRODUCTS)
        supplier = rng.choice(SUPPLIERS)
        quantity = rng.randint(6, 240)
        unit_price = round(base_price * rng.uniform(0.92, 1.12), 2)
        total = round(quantity * unit_price, 2)
        order_date = start + timedelta(days=rng.randrange(175))

        order_id = f"PO-2026-{1001 + index}"
        if rng.random() < 0.15:
            order_id = f" {order_id.lower()} "

        records.append(
            {
                "  order id  ": _maybe_missing(order_id, rng, 0.01),
                "order_date": _maybe_missing(order_date.strftime(rng.choice(DATE_FORMATS)), rng, 0.04),
                "supplier name": _maybe_missing(_case_variant(supplier, rng), rng, 0.03),
                "product name": _maybe_missing(product, rng, 0.02),
                "category": _maybe_missing(_case_variant(category, rng), rng, 0.02),
                "qty": _maybe_missing(quantity, rng, 0.03),
                "unit price ($)": _maybe_missing(_currency(unit_price, rng), rng, 0.04),
                "total amount": _maybe_missing(_currency(total, rng), rng, 0.04),
                "delivery_status": _maybe_missing(_case_variant(rng.choice(STATUSES), rng), rng, 0.05),
                "warehouse region": _maybe_missing(_case_variant(rng.choice(REGIONS), rng), rng, 0.03),
            }
        )

    # Inject exact duplicates, the kind created by copy-paste exports.
    duplicate_count = max(6, rows // 12)
    duplicate_indexes = rng.sample(range(rows), duplicate_count)
    records.extend(records[i].copy() for i in duplicate_indexes)
    rng.shuffle(records)

    return pd.DataFrame(records, columns=COLUMNS)


if __name__ == "__main__":
    data = generate_sample_data()
    print(f"Generated {len(data)} messy rows x {len(data.columns)} columns")
    print(data.head(8).to_string())

# Backend/entity_configs/Supplier_Item.py

# ── Column aliases ────────────────────────────────────────────────────────────
# Maps KNIME column names → internal names your loader scripts use
# Columns not listed here pass through as-is (loader ignores unknown columns)

COLUMN_ALIASES = {
    "Item Number":  "Item Code",
    "Supplier":     "Supplier Code",
}

# ── Default values ────────────────────────────────────────────────────────────
# Injected into every row after alias mapping
# Used for fields KNIME doesn't send but your loader requires

DEFAULTS = {
    "Domain Code": "10USA",   # hardcoded for now
}

# ── Optional fields ───────────────────────────────────────────────────────────
# These columns are accepted from KNIME but not required by the loader
# If missing or empty, they pass through as "" without raising any error

OPTIONAL_FIELDS = [
    "Supplier Lead Time",
    "Use SO Reduction Price",
    "SO Price Reduction",
    "Currency",
    "Quote Qty",
    "Price List",
    "Manufacturer",
    "Manufacturer Item",
    "Comment",
]
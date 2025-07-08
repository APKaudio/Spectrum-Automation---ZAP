import numpy as np
import itertools
import csv

# --- TRANSMITTER FREQUENCIES ---

# First 50 TXs spaced 25 MHz apart starting at 500 MHz
tx_freqs_1 = np.arange(500, 500 + 25 * 50, 25)

# Next 50 TXs spaced 10 MHz apart, starting from the last +10
tx_freqs_2 = np.arange(tx_freqs_1[-1] + 10, tx_freqs_1[-1] + 10 + 10 * 50, 10)

tx_freqs = np.concatenate((tx_freqs_1, tx_freqs_2))

# --- COORDINATION RULES ---

# Allowed RF band
band_min = 470
band_max = 698

# Forbidden range(s) (e.g., 600 MHz repack)
forbidden_bands = [
    (620, 650)
]

def is_in_forbidden_band(freq):
    return any(fmin <= freq <= fmax for fmin, fmax in forbidden_bands)

# --- CALCULATE IMD PRODUCTS ---

imd_records = []

for f1, f2 in itertools.combinations(tx_freqs, 2):
    imd1 = 2 * f1 - f2
    imd2 = 2 * f2 - f1

    if band_min <= imd1 <= band_max and not is_in_forbidden_band(imd1):
        imd_records.append({
            "IMD_Frequency_MHz": round(imd1, 3),
            "Formula": f"2*{f1} - {f2}",
            "Type": "2f1 - f2"
        })

    if band_min <= imd2 <= band_max and not is_in_forbidden_band(imd2):
        imd_records.append({
            "IMD_Frequency_MHz": round(imd2, 3),
            "Formula": f"2*{f2} - {f1}",
            "Type": "2f2 - f1"
        })

# --- EXPORT TO CSV ---

output_file = "imd_products.csv"

with open(output_file, "w", newline="") as csvfile:
    fieldnames = ["IMD_Frequency_MHz", "Formula", "Type"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    writer.writeheader()
    for row in imd_records:
        writer.writerow(row)

print(f"Exported {len(imd_records)} IMD products to {output_file}")

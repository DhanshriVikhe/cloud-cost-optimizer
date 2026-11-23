#!/usr/bin/env python3
"""
Azure Pricing Automation Tool
Reads VM inventory files (RVTools, CSV, etc.), right-sizes VMs,
and outputs an Excel workbook with Azure pricing across multiple tiers.

Usage:
  python azure_pricing_tool.py <input.xlsx>                        # Default N-1 (CPU>8, MEM>32)
  python azure_pricing_tool.py <input.xlsx> --like-for-like        # No downsizing
  python azure_pricing_tool.py <input.xlsx> --downsize 8 24        # Custom thresholds
"""

import json
import math
import sys
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

# =============================================================================
# AZURE VM PRICING DATABASE
# Prices: USD/month, East US region
# Source: Microsoft Azure pricing pages (March 2026)
#   Linux/BYOL: https://azure.microsoft.com/en-us/pricing/details/virtual-machines/linux/
#   Windows LI: https://azure.microsoft.com/en-us/pricing/details/virtual-machines/windows/
# Commitment type: Savings Plan (1 Year / 3 Year)
# BYOL (Azure Hybrid Benefit) = Linux rate (no OS license cost)
# Each entry stores MONTHLY prices directly from Microsoft.
# =============================================================================

# Reserved Instance 3yr prices (Linux, East US) — different from Savings Plan for D/E series
# Source: https://azure.microsoft.com/en-us/pricing/details/virtual-machines/linux/
# Filter: Reserved Instances (1 & 3 year), East US, Ubuntu
RI_3YR_PRICES = {
    "Standard_B2as_v2": 20.8611, "Standard_B4as_v2": 41.7222, "Standard_B8as_v2": 83.4444,
    "Standard_B16as_v2": 166.8889, "Standard_B32as_v2": 333.7778,
    "Standard_B2als_v2": 10.4167, "Standard_B4als_v2": 36.9444, "Standard_B8als_v2": 73.9167,
    "Standard_B16als_v2": 147.7778, "Standard_B32als_v2": 295.5556,
    "Standard_B2s_v2": 23.0833, "Standard_B4s_v2": 46.1667, "Standard_B8s_v2": 92.3056,
    "Standard_B16s_v2": 184.6389, "Standard_B32s_v2": 369.2778,
    "Standard_D2as_v5": 23.8611, "Standard_D4as_v5": 47.7222, "Standard_D8as_v5": 95.4167,
    "Standard_D16as_v5": 190.8611, "Standard_D32as_v5": 381.6944, "Standard_D48as_v5": 572.5556,
    "Standard_D2s_v5": 25.9167, "Standard_D4s_v5": 51.8611, "Standard_D8s_v5": 103.7222,
    "Standard_D16s_v5": 207.4444, "Standard_D32s_v5": 414.8611,
    "Standard_E2as_v5": 31.3333, "Standard_E4as_v5": 62.6944, "Standard_E8as_v5": 125.3889,
    "Standard_E16as_v5": 250.7778, "Standard_E20as_v5": 313.4722, "Standard_E32as_v5": 501.5278,
    "Standard_E48as_v5": 752.3056, "Standard_E64as_v5": 1003.0833, "Standard_E96as_v5": 1504.6111,
}

AZURE_VM_PRICING = [
    # ===== B-series AMD (Bas v2) - BURSTABLE =====
    # RHEL/SUSE prices sourced directly from Azure RHEL/SUSE pricing pages (East US, Savings Plan)
    {"sku": "Standard_B2as_v2",  "series": "B", "vcpus": 2,  "mem_gib": 8,   "linux_payg": 54.90,   "win_payg": 61.61,    "linux_1yr": 36.78,   "win_1yr": 42.62,   "linux_3yr": 24.70,   "win_3yr": 30.54,   "rhel_payg": 75.92,   "rhel_1yr": 57.80,   "rhel_3yr": 45.73, "suse_payg": 102.35, "suse_1yr": 84.23, "suse_3yr": 72.15},
    {"sku": "Standard_B4as_v2",  "series": "B", "vcpus": 4,  "mem_gib": 16,  "linux_payg": 109.50,  "win_payg": 123.37,   "linux_1yr": 73.37,   "win_1yr": 85.05,   "linux_3yr": 49.28,   "win_3yr": 60.96,   "rhel_payg": 151.55,  "rhel_1yr": 115.41,  "rhel_3yr": 91.32, "suse_payg": 200.75, "suse_1yr": 164.62, "suse_3yr": 140.53},
    {"sku": "Standard_B8as_v2",  "series": "B", "vcpus": 8,  "mem_gib": 32,  "linux_payg": 219.73,  "win_payg": 246.74,   "linux_1yr": 147.22,  "win_1yr": 170.58,  "linux_3yr": 98.88,   "win_3yr": 122.24,  "rhel_payg": 303.83,  "rhel_1yr": 231.32,  "rhel_3yr": 182.97, "suse_payg": 329.23, "suse_1yr": 256.72, "suse_3yr": 208.38},
    {"sku": "Standard_B16as_v2", "series": "B", "vcpus": 16, "mem_gib": 64,  "linux_payg": 439.46,  "win_payg": 492.75,   "linux_1yr": 294.44,  "win_1yr": 341.16,  "linux_3yr": 197.76,  "win_3yr": 244.48,  "rhel_payg": 565.60,  "rhel_1yr": 420.58,  "rhel_3yr": 323.90, "suse_payg": 548.96, "suse_1yr": 403.94, "suse_3yr": 307.26},
    {"sku": "Standard_B32as_v2", "series": "B", "vcpus": 32, "mem_gib": 128, "linux_payg": 878.19,  "win_payg": 985.50,   "linux_1yr": 588.39,  "win_1yr": 681.83,  "linux_3yr": 395.19,  "win_3yr": 488.63,  "rhel_payg": 1130.48, "rhel_1yr": 840.68,  "rhel_3yr": 647.47, "suse_payg": 987.69, "suse_1yr": 697.89, "suse_3yr": 504.69},

    # ===== B-series AMD low-memory (Bals v2) =====
    {"sku": "Standard_B2als_v2",  "series": "B", "vcpus": 2,  "mem_gib": 4,   "linux_payg": 27.45,   "win_payg": 34.16,    "linux_1yr": 18.39,   "win_1yr": 24.23,   "linux_3yr": 12.36,   "win_3yr": 18.20,   "rhel_payg": 48.47,   "rhel_1yr": 39.41,   "rhel_3yr": 33.38, "suse_payg": 74.90, "suse_1yr": 65.84, "suse_3yr": 59.80},
    {"sku": "Standard_B4als_v2",  "series": "B", "vcpus": 4,  "mem_gib": 8,   "linux_payg": 97.09,   "win_payg": 110.96,   "linux_1yr": 65.05,   "win_1yr": 76.73,   "linux_3yr": 43.69,   "win_3yr": 55.37,   "rhel_payg": 139.14,  "rhel_1yr": 107.10,  "rhel_3yr": 85.74, "suse_payg": 188.34, "suse_1yr": 156.30, "suse_3yr": 134.94},
    {"sku": "Standard_B8als_v2",  "series": "B", "vcpus": 8,  "mem_gib": 16,  "linux_payg": 194.18,  "win_payg": 221.19,   "linux_1yr": 130.10,  "win_1yr": 153.46,  "linux_3yr": 87.38,   "win_3yr": 110.74,  "rhel_payg": 278.28,  "rhel_1yr": 214.20,  "rhel_3yr": 171.48, "suse_payg": 303.68, "suse_1yr": 239.60, "suse_3yr": 196.88},
    {"sku": "Standard_B16als_v2", "series": "B", "vcpus": 16, "mem_gib": 32,  "linux_payg": 389.09,  "win_payg": 442.38,   "linux_1yr": 260.69,  "win_1yr": 307.41,  "linux_3yr": 175.09,  "win_3yr": 221.81,  "rhel_payg": 515.23,  "rhel_1yr": 386.83,  "rhel_3yr": 301.23, "suse_payg": 498.59, "suse_1yr": 370.19, "suse_3yr": 284.59},
    {"sku": "Standard_B32als_v2", "series": "B", "vcpus": 32, "mem_gib": 64,  "linux_payg": 777.45,  "win_payg": 884.76,   "linux_1yr": 520.89,  "win_1yr": 614.33,  "linux_3yr": 349.85,  "win_3yr": 443.29,  "rhel_payg": 1029.74, "rhel_1yr": 773.18,  "rhel_3yr": 602.14, "suse_payg": 886.95, "suse_1yr": 630.39, "suse_3yr": 459.35},

    # ===== B-series Intel (Bs v2) =====
    {"sku": "Standard_B2s_v2",   "series": "B", "vcpus": 2,  "mem_gib": 8,   "linux_payg": 60.74,   "win_payg": 67.45,    "linux_1yr": 40.83,   "win_1yr": 46.53,   "linux_3yr": 29.14,   "win_3yr": 33.17,   "rhel_payg": 81.76,   "rhel_1yr": 61.71,   "rhel_3yr": 48.36, "suse_payg": 108.19, "suse_1yr": 88.14, "suse_3yr": 74.78},
    {"sku": "Standard_B4s_v2",   "series": "B", "vcpus": 4,  "mem_gib": 16,  "linux_payg": 121.54,  "win_payg": 135.05,   "linux_1yr": 81.48,   "win_1yr": 92.87,   "linux_3yr": 58.09,   "win_3yr": 66.21,   "rhel_payg": 163.23,  "rhel_1yr": 123.24,  "rhel_3yr": 96.58, "suse_payg": 212.43, "suse_1yr": 172.44, "suse_3yr": 145.78},
    {"sku": "Standard_B8s_v2",   "series": "B", "vcpus": 8,  "mem_gib": 32,  "linux_payg": 243.09,  "win_payg": 270.10,   "linux_1yr": 162.97,  "win_1yr": 186.23,  "linux_3yr": 116.17,  "win_3yr": 132.75,  "rhel_payg": 327.19,  "rhel_1yr": 246.97,  "rhel_3yr": 193.49, "suse_payg": 352.59, "suse_1yr": 272.37, "suse_3yr": 218.89},
    {"sku": "Standard_B16s_v2",  "series": "B", "vcpus": 16, "mem_gib": 64,  "linux_payg": 485.52,  "win_payg": 539.47,   "linux_1yr": 325.94,  "win_1yr": 372.46,  "linux_3yr": 232.35,  "win_3yr": 265.50,  "rhel_payg": 612.32,  "rhel_1yr": 451.88,  "rhel_3yr": 344.93, "suse_payg": 595.68, "suse_1yr": 435.24, "suse_3yr": 328.28},
    {"sku": "Standard_B32s_v2",  "series": "B", "vcpus": 32, "mem_gib": 128, "linux_payg": 971.05,  "win_payg": 1078.94,  "linux_1yr": 651.88,  "win_1yr": 744.43,  "linux_3yr": 464.70,  "win_3yr": 530.67,  "rhel_payg": 1223.92, "rhel_1yr": 903.28,  "rhel_3yr": 689.52, "suse_payg": 1081.13, "suse_1yr": 760.49, "suse_3yr": 546.73},

    # ===== D-series v5 AMD (Das v5) - GENERAL PURPOSE =====
    {"sku": "Standard_D2as_v5",  "series": "D", "vcpus": 2,  "mem_gib": 8,   "linux_payg": 62.78,   "win_payg": 129.94,   "linux_1yr": 42.89,   "win_1yr": 110.05,  "linux_3yr": 28.80,   "win_3yr": 95.96,   "rhel_payg": 83.80,   "rhel_1yr": 63.91,   "rhel_3yr": 49.82, "suse_payg": 110.23, "suse_1yr": 90.34, "suse_3yr": 76.25},
    {"sku": "Standard_D4as_v5",  "series": "D", "vcpus": 4,  "mem_gib": 16,  "linux_payg": 125.56,  "win_payg": 259.88,   "linux_1yr": 85.77,   "win_1yr": 220.09,  "linux_3yr": 57.60,   "win_3yr": 191.92,  "rhel_payg": 167.61,  "rhel_1yr": 127.82,  "rhel_3yr": 99.65, "suse_payg": 216.81, "suse_1yr": 177.02, "suse_3yr": 148.85},
    {"sku": "Standard_D8as_v5",  "series": "D", "vcpus": 8,  "mem_gib": 32,  "linux_payg": 251.12,  "win_payg": 519.76,   "linux_1yr": 171.54,  "win_1yr": 440.18,  "linux_3yr": 115.19,  "win_3yr": 383.83,  "rhel_payg": 335.22,  "rhel_1yr": 255.64,  "rhel_3yr": 199.28, "suse_payg": 360.62, "suse_1yr": 281.04, "suse_3yr": 224.69},
    {"sku": "Standard_D16as_v5", "series": "D", "vcpus": 16, "mem_gib": 64,  "linux_payg": 502.24,  "win_payg": 1039.52,  "linux_1yr": 343.08,  "win_1yr": 880.36,  "linux_3yr": 230.38,  "win_3yr": 767.66,  "rhel_payg": 628.38,  "rhel_1yr": 469.22,  "rhel_3yr": 356.52, "suse_payg": 611.74, "suse_1yr": 452.58, "suse_3yr": 339.88},
    {"sku": "Standard_D32as_v5", "series": "D", "vcpus": 32, "mem_gib": 128, "linux_payg": 1004.48, "win_payg": 2079.04,  "linux_1yr": 686.16,  "win_1yr": 1760.72, "linux_3yr": 460.75,  "win_3yr": 1535.31, "rhel_payg": 1256.77, "rhel_1yr": 938.45,  "rhel_3yr": 713.04, "suse_payg": 1113.98, "suse_1yr": 795.66, "suse_3yr": 570.25},

    # ===== D-series v5 Intel (Ds v5) =====
    {"sku": "Standard_D2s_v5",   "series": "D", "vcpus": 2,  "mem_gib": 8,   "linux_payg": 66.36,   "win_payg": 137.24,   "linux_1yr": 55.77,   "win_1yr": 115.22,  "linux_3yr": 48.18,   "win_3yr": 99.41,   "rhel_payg": 91.10,   "rhel_1yr": 69.09,   "rhel_3yr": 53.28, "suse_payg": 117.53, "suse_1yr": 95.51, "suse_3yr": 79.7},
    {"sku": "Standard_D4s_v5",   "series": "D", "vcpus": 4,  "mem_gib": 16,  "linux_payg": 132.72,  "win_payg": 274.48,   "linux_1yr": 111.53,  "win_1yr": 230.44,  "linux_3yr": 96.36,   "win_3yr": 198.82,  "rhel_payg": 182.21,  "rhel_1yr": 138.17,  "rhel_3yr": 106.55, "suse_payg": 231.41, "suse_1yr": 187.37, "suse_3yr": 155.75},
    {"sku": "Standard_D8s_v5",   "series": "D", "vcpus": 8,  "mem_gib": 32,  "linux_payg": 265.44,  "win_payg": 548.96,   "linux_1yr": 223.07,  "win_1yr": 460.89,  "linux_3yr": 192.72,  "win_3yr": 397.65,  "rhel_payg": 364.42,  "rhel_1yr": 276.34,  "rhel_3yr": 213.10, "suse_payg": 389.82, "suse_1yr": 301.75, "suse_3yr": 238.51},
    {"sku": "Standard_D16s_v5",  "series": "D", "vcpus": 16, "mem_gib": 64,  "linux_payg": 530.88,  "win_payg": 1097.92,  "linux_1yr": 446.13,  "win_1yr": 921.76,  "linux_3yr": 385.44,  "win_3yr": 795.28,  "rhel_payg": 686.78,  "rhel_1yr": 510.63,  "rhel_3yr": 384.15, "suse_payg": 670.14, "suse_1yr": 493.98, "suse_3yr": 367.5},
    {"sku": "Standard_D32s_v5",  "series": "D", "vcpus": 32, "mem_gib": 128, "linux_payg": 1061.76, "win_payg": 2195.84,  "linux_1yr": 892.26,  "win_1yr": 1843.53, "linux_3yr": 770.88,  "win_3yr": 1590.58, "rhel_payg": 1373.57, "rhel_1yr": 1021.26, "rhel_3yr": 768.30, "suse_payg": 1230.78, "suse_1yr": 878.47, "suse_3yr": 625.52},

    # ===== E-series v5 AMD (Eas v5) - MEMORY OPTIMIZED =====
    {"sku": "Standard_E2as_v5",  "series": "E", "vcpus": 2,  "mem_gib": 16,  "linux_payg": 82.49,   "win_payg": 149.65,   "linux_1yr": 55.94,   "win_1yr": 123.10,  "linux_3yr": 39.05,   "win_3yr": 106.21,  "rhel_payg": 103.51,  "rhel_1yr": 76.96,   "rhel_3yr": 60.07, "suse_payg": 129.94, "suse_1yr": 103.39, "suse_3yr": 86.5},
    {"sku": "Standard_E4as_v5",  "series": "E", "vcpus": 4,  "mem_gib": 32,  "linux_payg": 164.98,  "win_payg": 299.30,   "linux_1yr": 111.87,  "win_1yr": 246.19,  "linux_3yr": 78.10,   "win_3yr": 212.42,  "rhel_payg": 207.03,  "rhel_1yr": 153.92,  "rhel_3yr": 120.15, "suse_payg": 256.23, "suse_1yr": 203.12, "suse_3yr": 169.35},
    {"sku": "Standard_E8as_v5",  "series": "E", "vcpus": 8,  "mem_gib": 64,  "linux_payg": 329.96,  "win_payg": 598.60,   "linux_1yr": 223.75,  "win_1yr": 492.39,  "linux_3yr": 156.21,  "win_3yr": 424.85,  "rhel_payg": 414.06,  "rhel_1yr": 307.84,  "rhel_3yr": 240.30, "suse_payg": 439.46, "suse_1yr": 333.25, "suse_3yr": 265.71},
    {"sku": "Standard_E16as_v5", "series": "E", "vcpus": 16, "mem_gib": 128, "linux_payg": 659.92,  "win_payg": 1197.20,  "linux_1yr": 447.49,  "win_1yr": 984.77,  "linux_3yr": 312.40,  "win_3yr": 849.68,  "rhel_payg": 786.06,  "rhel_1yr": 573.63,  "rhel_3yr": 438.55, "suse_payg": 769.42, "suse_1yr": 556.99, "suse_3yr": 421.9},
    {"sku": "Standard_E32as_v5", "series": "E", "vcpus": 32, "mem_gib": 256, "linux_payg": 1319.84, "win_payg": 2394.40,  "linux_1yr": 894.98,  "win_1yr": 1969.54, "linux_3yr": 624.81,  "win_3yr": 1699.37, "rhel_payg": 1572.13, "rhel_1yr": 1147.27, "rhel_3yr": 877.10, "suse_payg": 1429.34, "suse_1yr": 1004.48, "suse_3yr": 734.31},
    {"sku": "Standard_E48as_v5", "series": "E", "vcpus": 48, "mem_gib": 384, "linux_payg": 1979.76, "win_payg": 3591.60,  "linux_1yr": 1342.56, "win_1yr": 2954.40, "linux_3yr": 937.20,  "win_3yr": 2549.04, "rhel_payg": 2484.24, "rhel_1yr": 1847.04, "rhel_3yr": 1441.68, "suse_payg": 3118.56, "suse_1yr": 2481.36, "suse_3yr": 2076.00},

    # ===== E-series v5 Intel (Es v5) =====
    {"sku": "Standard_E2s_v5",   "series": "E", "vcpus": 2,  "mem_gib": 16,  "linux_payg": 76.65,   "win_payg": 159.14,   "linux_1yr": 62.05,   "win_1yr": 129.36,  "linux_3yr": 53.46,   "win_3yr": 110.70,  "rhel_payg": 113.00,  "rhel_1yr": 83.22,   "rhel_3yr": 64.56, "suse_payg": 139.43, "suse_1yr": 109.65, "suse_3yr": 90.99},
    {"sku": "Standard_E4s_v5",   "series": "E", "vcpus": 4,  "mem_gib": 32,  "linux_payg": 153.30,  "win_payg": 318.28,   "linux_1yr": 124.10,  "win_1yr": 258.71,  "linux_3yr": 106.92,  "win_3yr": 221.39,  "rhel_payg": 226.01,  "rhel_1yr": 166.44,  "rhel_3yr": 129.12, "suse_payg": 275.21, "suse_1yr": 215.64, "suse_3yr": 178.32},
    {"sku": "Standard_E8s_v5",   "series": "E", "vcpus": 8,  "mem_gib": 64,  "linux_payg": 306.60,  "win_payg": 636.56,   "linux_1yr": 248.20,  "win_1yr": 517.42,  "linux_3yr": 213.84,  "win_3yr": 442.77,  "rhel_payg": 452.02,  "rhel_1yr": 332.88,  "rhel_3yr": 258.23, "suse_payg": 477.42, "suse_1yr": 358.28, "suse_3yr": 283.63},
    {"sku": "Standard_E16s_v5",  "series": "E", "vcpus": 16, "mem_gib": 128, "linux_payg": 613.20,  "win_payg": 1273.12,  "linux_1yr": 496.40,  "win_1yr": 1034.86, "linux_3yr": 427.68,  "win_3yr": 885.56,  "rhel_payg": 861.98,  "rhel_1yr": 623.72,  "rhel_3yr": 474.42, "suse_payg": 845.34, "suse_1yr": 607.08, "suse_3yr": 457.78},
    {"sku": "Standard_E32s_v5",  "series": "E", "vcpus": 32, "mem_gib": 256, "linux_payg": 1226.40, "win_payg": 2546.24,  "linux_1yr": 992.80,  "win_1yr": 2069.71, "linux_3yr": 855.36,  "win_3yr": 1771.10, "rhel_payg": 1723.97, "rhel_1yr": 1247.44, "rhel_3yr": 948.83, "suse_payg": 1581.18, "suse_1yr": 1104.65, "suse_3yr": 806.04},

    # ===== F-series v6 AMD (Fas v6) - COMPUTE OPTIMIZED =====
    {"sku": "Standard_F2as_v6",  "series": "F", "vcpus": 2,  "mem_gib": 8,   "linux_payg": 100.73,  "win_payg": 167.17,   "linux_1yr": 79.66,   "win_1yr": 136.17,  "linux_3yr": 66.14,   "win_3yr": 113.16,  "rhel_payg": 121.03,  "rhel_1yr": 90.03,   "rhel_3yr": 67.03, "suse_payg": 147.46, "suse_1yr": 116.46, "suse_3yr": 93.45},
    {"sku": "Standard_F4as_v6",  "series": "F", "vcpus": 4,  "mem_gib": 16,  "linux_payg": 200.72,  "win_payg": 333.61,   "linux_1yr": 158.82,  "win_1yr": 271.83,  "linux_3yr": 131.92,  "win_3yr": 225.99,  "rhel_payg": 241.34,  "rhel_1yr": 179.56,  "rhel_3yr": 133.72, "suse_payg": 290.54, "suse_1yr": 228.76, "suse_3yr": 182.92},
    {"sku": "Standard_F8as_v6",  "series": "F", "vcpus": 8,  "mem_gib": 32,  "linux_payg": 401.44,  "win_payg": 667.22,   "linux_1yr": 317.64,  "win_1yr": 543.66,  "linux_3yr": 263.84,  "win_3yr": 451.99,  "rhel_payg": 482.68,  "rhel_1yr": 359.12,  "rhel_3yr": 267.44, "suse_payg": 508.08, "suse_1yr": 384.52, "suse_3yr": 292.85},
    {"sku": "Standard_F16as_v6", "series": "F", "vcpus": 16, "mem_gib": 64,  "linux_payg": 802.88,  "win_payg": 1335.17,  "linux_1yr": 635.28,  "win_1yr": 1087.82, "linux_3yr": 527.68,  "win_3yr": 904.31,  "rhel_payg": 924.03,  "rhel_1yr": 676.69,  "rhel_3yr": 493.17, "suse_payg": 907.39, "suse_1yr": 660.04, "suse_3yr": 476.53},
    {"sku": "Standard_F32as_v6", "series": "F", "vcpus": 32, "mem_gib": 128, "linux_payg": 1605.76, "win_payg": 2670.34,  "linux_1yr": 1270.56, "win_1yr": 2175.65, "linux_3yr": 1055.36, "win_3yr": 1808.62, "rhel_payg": 1848.07, "rhel_1yr": 1353.38, "rhel_3yr": 986.35, "suse_payg": 1705.28, "suse_1yr": 1210.59, "suse_3yr": 843.56},
]

# =============================================================================
# RHEL / SUSE LICENSE PREMIUM TABLES
# Source: Microsoft Azure pricing pages (East US, no AHB, Savings Plan)
# Premium = SLES/RHEL full price minus base Linux price, averaged across B/D/E
# Added on top of the base Linux rate for LI columns only.
# BYOL columns stay at base Linux rate (customer brings own license).
# RHEL premium is a flat fee by vCPU (consistent across all series).
# SUSE premium varies slightly but averaged per vCPU tier.
# =============================================================================

RHEL_PREMIUM = {
    # vCPUs: (payg_premium, 1yr_sp_premium, 3yr_sp_premium)
    # RHEL premium is FLAT across all commitment tiers (same for PAYG/1yr/3yr)
    # Verified from Azure RHEL pricing page vs Linux pricing page (East US, B/D/E series)
    2:  (21.02, 21.02, 21.02),
    4:  (42.05, 42.05, 42.05),
    8:  (84.10, 84.10, 84.10),
    16: (126.14, 126.14, 126.14),
    32: (252.29, 252.29, 252.29),
}

SUSE_PREMIUM = {
    # vCPUs: (payg_premium, 1yr_sp_premium, 3yr_sp_premium)
    # SUSE premium is FLAT across all commitment tiers (same for PAYG/1yr/3yr)
    # Verified from Azure SUSE pricing page vs Linux pricing page (East US, B/D/E series)
    # Flat $109.50/mo for 8+ vCPUs
    2:  (47.45, 47.45, 47.45),
    4:  (91.25, 91.25, 91.25),
    8:  (109.50, 109.50, 109.50),
    16: (109.50, 109.50, 109.50),
    32: (109.50, 109.50, 109.50),
}

def get_license_premium(vcpus, distro):
    """Get the license premium for RHEL or SUSE based on Azure vCPU count.
    Returns (payg_premium, 1yr_premium, 3yr_premium) or (0,0,0) for free Linux."""
    if distro == "rhel":
        table = RHEL_PREMIUM
    elif distro == "suse":
        table = SUSE_PREMIUM
    else:
        return (0.0, 0.0, 0.0)

    tiers = sorted(table.keys())
    for t in tiers:
        if vcpus <= t:
            return table[t]
    return table[tiers[-1]]

def detect_linux_distro(os_detail):
    """Classify Linux distro from os_detail string.
    Returns: 'rhel', 'suse', 'oracle', 'free' """
    od = str(os_detail).lower()
    if "red hat" in od or "rhel" in od or "centos" in od:
        return "rhel"
    elif "suse" in od or "sles" in od:
        return "suse"
    elif "oracle linux" in od:
        return "oracle"
    else:
        return "free"

# Azure Managed Disk Pricing (Standard SSD, East US, monthly)
MANAGED_DISK_PRICING = [
    {"tier": "E1",  "size_gib": 4,    "monthly": 0.30},
    {"tier": "E2",  "size_gib": 8,    "monthly": 0.60},
    {"tier": "E3",  "size_gib": 16,   "monthly": 1.20},
    {"tier": "E4",  "size_gib": 32,   "monthly": 2.40},
    {"tier": "E6",  "size_gib": 64,   "monthly": 4.80},
    {"tier": "E10", "size_gib": 128,  "monthly": 9.60},
    {"tier": "E15", "size_gib": 256,  "monthly": 19.20},
    {"tier": "E20", "size_gib": 512,  "monthly": 38.40},
    {"tier": "E30", "size_gib": 1024, "monthly": 76.80},
    {"tier": "E40", "size_gib": 2048, "monthly": 153.60},
    {"tier": "E50", "size_gib": 4096, "monthly": 307.20},
    {"tier": "E60", "size_gib": 8192, "monthly": 614.40},
    {"tier": "E70", "size_gib": 16384,"monthly": 1228.80},
    {"tier": "E80", "size_gib": 32767,"monthly": 2457.60},
]

HOURS_PER_MONTH = 730

# =============================================================================
# CLI FLAGS (set by main())
# =============================================================================

LIKE_FOR_LIKE = False
DOWNSIZE_CPU_THRESHOLD = 8
DOWNSIZE_MEM_THRESHOLD = 32
PREFER_HIGHER_RAM = False
B_SERIES_ALL = True  # If False, B-series only for non-prod

# =============================================================================
# VM CLASSIFICATION
# =============================================================================

def detect_os_type(os_config, os_tools):
    """Determine if Windows or Linux based on OS fields."""
    os_str = (str(os_tools) + " " + str(os_config)).lower()
    if "windows" in os_str:
        return "Windows"
    return "Linux"

def is_desktop_os(os_config, os_tools):
    """Detect desktop OS (Windows 10, 8, 7, XP, Vista, 11) — these should be
    priced as Azure Virtual Desktop, not regular compute VMs.
    
    Prioritizes os_tools (VMware Tools reports actual running OS) over os_config
    (template setting, can be wrong). If os_tools says Server, it's not desktop."""
    tools_str = str(os_tools).lower()
    config_str = str(os_config).lower()
    
    # If VMware Tools reports Server OS, it's NOT desktop regardless of config
    if "server" in tools_str:
        return False
    
    desktop_patterns = ["windows 10", "windows 11", "windows 7", "windows 8",
                        "windows xp", "windows vista"]
    
    # Check os_tools first (more reliable), then os_config
    if tools_str and tools_str != "none":
        return any(p in tools_str for p in desktop_patterns)
    return any(p in config_str for p in desktop_patterns)

def detect_sql(vm_name, annotation, folder, cluster, os_type):
    """Detect SQL workloads from multiple fields."""
    check_str = f"{vm_name} {annotation} {folder} {cluster}".lower()
    sql_patterns = ["sql", "database", "mssql"]
    if any(p in check_str for p in sql_patterns):
        if os_type == "Windows":
            return True
    return False

def detect_nonprod(vm_name, folder, cluster, source_file=""):
    """Detect non-production workloads.
    
    Primary method: Source File (vCenter server name).
    vCenters with '-PROD' in the name are production; others are non-prod.
    This matches Concierto's classification (1,279 non-prod vs our 1,280).
    
    Fallback: VM name, folder, and cluster keyword matching.
    """
    # Primary: vCenter source file classification
    if source_file:
        if "-PROD" in source_file.upper():
            # From a production vCenter — but still check name patterns
            # (a dev VM can live on a prod vCenter)
            vm_lower = vm_name.lower()
            nonprod_name_patterns = [
                "dev", "tst", "test", "uat", "stg", "staging",
                "lab", "nprod", "nonprod", "sandbox", "demo", "poc", "tmp"
            ]
            if any(p in vm_lower for p in nonprod_name_patterns):
                return True
            return False
        else:
            # From a non-production vCenter — tag as non-prod
            return True

    # Fallback: keyword-based detection (when Source File not available)
    vm_lower = vm_name.lower()
    folder_lower = folder.lower()
    cluster_lower = cluster.lower()

    nonprod_name_patterns = [
        "dev", "tst", "test", "uat", "stg", "staging",
        "lab", "nprod", "nonprod", "sandbox", "demo", "poc", "tmp"
    ]
    if any(p in vm_lower for p in nonprod_name_patterns):
        return True
    if "non prod" in folder_lower or "nonprod" in folder_lower or "non_prod" in folder_lower:
        return True
    if "nprod" in cluster_lower:
        return True
    return False

# =============================================================================
# RIGHT SIZING LOGIC
# =============================================================================

def apply_downsizing(vcpus, mem_gib):
    """
    Downsizes VMs based on configurable thresholds.
    --like-for-like: no downsizing
    --downsize CPU MEM: custom thresholds (default 8 32)
    If CPU > threshold OR memory > threshold, step down one size tier.
    """
    if LIKE_FOR_LIKE:
        return vcpus, mem_gib

    size_tiers_cpu = [1, 2, 4, 8, 16, 32, 48, 64, 96]
    size_tiers_mem = [1, 2, 4, 8, 16, 24, 32, 64, 128, 256, 512]

    new_cpu = vcpus
    new_mem = mem_gib

    if vcpus > DOWNSIZE_CPU_THRESHOLD:
        for i, tier in enumerate(size_tiers_cpu):
            if tier >= vcpus:
                new_cpu = size_tiers_cpu[max(0, i - 1)]
                break
        else:
            new_cpu = size_tiers_cpu[-2]

    if mem_gib > DOWNSIZE_MEM_THRESHOLD:
        for i, tier in enumerate(size_tiers_mem):
            if tier >= mem_gib:
                new_mem = size_tiers_mem[max(0, i - 1)]
                break
        else:
            new_mem = size_tiers_mem[-2]

    return new_cpu, new_mem

def find_cheapest_vm(target_cpu, target_mem_gib, is_windows, is_nonprod, linux_distro="free", prefer_higher_ram=False):
    """Find the cheapest Azure VM SKU that meets or exceeds the target specs.
    
    Compares by total License Included price so that license premiums
    (Windows, RHEL, SUSE) factor into which SKU is actually cheapest.
    Uses actual per-SKU prices from Microsoft pricing pages.
    
    If prefer_higher_ram=True, among VMs with the same vCPU count, prefer
    the higher-memory variant (e.g. B4as v2 over B4als v2) when both fit.
    This matches SMC-style sizing that favors memory headroom.
    """
    candidates = []
    for vm in AZURE_VM_PRICING:
        if vm["vcpus"] >= target_cpu and vm["mem_gib"] >= target_mem_gib:
            # When B_SERIES_ALL is False, restrict B-series to non-prod only
            if not B_SERIES_ALL and not is_nonprod and vm["series"] == "B":
                continue
            # When prefer_higher_ram is set, skip low-memory "als" variants
            # so the tool picks standard "as" variants with more RAM headroom
            if prefer_higher_ram and "als" in vm["sku"].lower():
                continue

            if is_windows:
                compare_price = vm["win_payg"]
            elif linux_distro == "rhel":
                compare_price = vm["rhel_payg"]
            elif linux_distro == "suse":
                compare_price = vm["suse_payg"]
            else:
                compare_price = vm["linux_payg"]

            candidates.append((compare_price, vm))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]

def find_disk_tier(total_disk_gib):
    """Find the appropriate Standard SSD managed disk tier."""
    for disk in MANAGED_DISK_PRICING:
        if disk["size_gib"] >= total_disk_gib:
            return disk
    return MANAGED_DISK_PRICING[-1]

# =============================================================================
# RVTOOLS PARSER
# =============================================================================

def parse_rvtools(filepath):
    """Parse an RVTools export and extract VM details."""
    wb = load_workbook(filepath, read_only=True)

    print(f"  Sheets found: {len(wb.sheetnames)}")
    for s in wb.sheetnames:
        print(f"    - {s}")

    # Always read vInfo first to show the full picture
    total_vinfo = 0
    powered_on = 0
    powered_off = 0
    templates = 0
    if "vInfo" in wb.sheetnames:
        ws_info = wb["vInfo"]
        info_rows = list(ws_info.iter_rows(values_only=True))
        info_header = info_rows[0]

        # Build column index for vInfo — handle prefixed names
        info_col = {}
        for i, col in enumerate(info_header):
            if col:
                cname = str(col).strip()
                info_col[cname] = i
                # Also map stripped prefix version
                for prefix in ["vInfo", "vinfo"]:
                    if cname.startswith(prefix):
                        info_col[cname[len(prefix):]] = i

        # Find Powerstate and Template columns
        ps_col = info_col.get("Powerstate", info_col.get("Power State", info_col.get("PowerState", 1)))
        tmpl_col = info_col.get("Template", 2)

        for row in info_rows[1:]:
            if row[0] is None:
                continue
            total_vinfo += 1
            ps = str(row[ps_col] or "").lower()
            tmpl = str(row[tmpl_col] or "").lower()
            if tmpl == "true":
                templates += 1
            elif "poweredoff" in ps.replace(" ", ""):
                powered_off += 1
            else:
                powered_on += 1

        print(f"\n  Full Inventory (vInfo sheet):")
        print(f"    Total VMs found:    {total_vinfo}")
        print(f"    Powered On:         {powered_on}")
        print(f"    Powered Off:        {powered_off}")
        print(f"    Templates:          {templates}")

    # Now read the sheet we'll actually use for pricing
    if "In-Scope" in wb.sheetnames:
        ws = wb["In-Scope"]
        rows = list(ws.iter_rows(values_only=True))
        header_row = None
        for i, row in enumerate(rows):
            if row and (row[0] == "VM" or str(row[0] or "").startswith("vInfo")):
                header_row = i
                break
        if header_row is None:
            raise ValueError("Could not find header row in In-Scope sheet")
        header = rows[header_row]
        data = rows[header_row + 1:]
        print(f"\n  Using 'In-Scope' sheet for pricing (pre-filtered to active VMs)")
    elif "vInfo" in wb.sheetnames:
        ws = wb["vInfo"]
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        info_col2 = {}
        for i, col in enumerate(header):
            if col:
                cname = str(col).strip()
                info_col2[cname] = i
                for prefix in ["vInfo", "vinfo"]:
                    if cname.startswith(prefix):
                        info_col2[cname[len(prefix):]] = i

        ps_col2 = info_col2.get("Powerstate", info_col2.get("Power State", info_col2.get("PowerState", 1)))
        tmpl_col2 = info_col2.get("Template", 2)
        os_tools_col = info_col2.get("OS according to the VMware Tools", info_col2.get("OS according to the VMware tools", 0))
        os_config_col = info_col2.get("OS according to the configuration file", info_col2.get("OS according to the Configuration File", 0))

        data = []
        skipped = 0
        for row in rows[1:]:
            if row[0] is None:
                continue
            ps = str(row[ps_col2] or "").lower()
            tmpl = str(row[tmpl_col2] or "").lower()
            os_str = str(row[os_tools_col] or "").lower()
            os_str2 = str(row[os_config_col] or "").lower()
            if tmpl == "true" or "poweredoff" in ps.replace(" ", "") or "photon" in os_str or "photon" in os_str2:
                skipped += 1
                continue
            data.append(row)
        print(f"\n  Using 'vInfo' sheet (filtered out {skipped} powered off/template/photon VMs)")
    else:
        raise ValueError("No In-Scope or vInfo sheet found")

    # Build column index map
    col_map = {}
    for i, col in enumerate(header):
        if col:
            col_map[str(col).strip()] = i

    # Normalize column names — some RVTools versions prefix with sheet name (e.g., "vInfoVMName" instead of "VM")
    # Map alternate/prefixed names to our standard names
    COLUMN_ALIASES = {
        "VM": ["VM", "VM Name", "Name", "vInfoVMName", "vInfoVM Name"],
        "CPUs": ["CPUs", "Num CPUs", "vInfoCPUs", "vInfoNum CPUs", "vCPU", "vCPUs"],
        "Memory": ["Memory", "Memory MB", "MemoryMB", "vInfoMemory", "vInfoMemory MB"],
        "Powerstate": ["Powerstate", "Power State", "vInfoPowerstate", "vInfoPower State"],
        "Template": ["Template", "vInfoTemplate"],
        "OS according to the VMware Tools": ["OS according to the VMware Tools", "vInfoOS according to the VMware Tools",
            "OS according to the VMware tools", "vInfoOS according to the VMware tools", "OS (VMware Tools)", "vInfoOS"],
        "OS according to the configuration file": ["OS according to the configuration file",
            "vInfoOS according to the configuration file", "OS according to the Configuration File",
            "vInfoOS according to the Configuration File", "OS (Config)", "vInfoConfig OS"],
        "Total disk capacity MiB": ["Total disk capacity MiB", "vInfoTotal disk capacity MiB",
            "Provisioned MiB", "vInfoProvisioned MiB", "Total Disk Capacity MiB"],
        "Annotation": ["Annotation", "vInfoAnnotation", "Notes"],
        "Folder": ["Folder", "vInfoFolder", "VM Folder"],
        "Cluster": ["Cluster", "vInfoCluster"],
        "Source File": ["Source File", "vInfoSource File", "Source", "vSource"],
        "Disks": ["Disks", "vInfoDisks", "vInfoNumVirtualDisks", "NumVirtualDisks"],
        "DNS Name": ["DNS Name", "vInfoDNS Name", "vInfoGuestHostName", "GuestHostName"],
        "Primary IP Address": ["Primary IP Address", "vInfoPrimary IP Address", "IP Address"],
        "EFI Secure boot": ["EFI Secure boot", "vInfoEFI Secure boot"],
    }

    # Build normalized map: find the first alias that exists in col_map
    norm = {}
    for standard_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in col_map:
                norm[standard_name] = col_map[alias]
                break

    # Also try fuzzy matching: strip "vInfo" prefix and common variations
    for raw_name, idx in col_map.items():
        # Strip known prefixes
        clean = raw_name
        for prefix in ["vInfo", "vinfo", "VINFO"]:
            if clean.startswith(prefix):
                clean = clean[len(prefix):]
                break
        # Check if this cleaned name matches any standard name (case-insensitive)
        for standard_name in COLUMN_ALIASES:
            if standard_name not in norm:
                if clean.lower() == standard_name.lower():
                    norm[standard_name] = idx
                elif clean.lower().replace(" ", "") == standard_name.lower().replace(" ", ""):
                    norm[standard_name] = idx

    # Required columns
    required = ["VM", "CPUs", "Memory"]
    for r in required:
        if r not in norm:
            raise ValueError(f"Required column '{r}' not found. Available columns: {list(col_map.keys())[:30]}\nTip: This might be a newer RVTools format. Please check column names.")

    # Print detected column mapping for transparency
    print(f"  Column mapping:")
    for std_name in ["VM", "CPUs", "Memory", "Powerstate", "OS according to the VMware Tools",
                     "OS according to the configuration file", "Total disk capacity MiB",
                     "Annotation", "Folder", "Cluster", "Source File"]:
        if std_name in norm:
            # Find original column name
            orig = [k for k, v in col_map.items() if v == norm[std_name]][0]
            mapped = f" (from '{orig}')" if orig != std_name else ""
            print(f"    {std_name}: col {norm[std_name]}{mapped}")
        else:
            print(f"    {std_name}: not found (optional)")

    vms = []
    for row in data:
        vm_name = row[norm["VM"]]
        if not vm_name:
            continue

        cpus = row[norm.get("CPUs", 0)]
        mem_mib = row[norm.get("Memory", 0)]
        if not cpus or not mem_mib:
            continue

        disk_mib = row[norm.get("Total disk capacity MiB", 0)] or 0 if "Total disk capacity MiB" in norm else 0
        os_config = str(row[norm.get("OS according to the configuration file", 0)] or "") if "OS according to the configuration file" in norm else ""
        os_tools = str(row[norm.get("OS according to the VMware Tools", 0)] or "") if "OS according to the VMware Tools" in norm else ""
        annotation = str(row[norm.get("Annotation", 0)] or "") if "Annotation" in norm else ""
        folder = str(row[norm.get("Folder", 0)] or "") if "Folder" in norm else ""
        cluster = str(row[norm.get("Cluster", 0)] or "") if "Cluster" in norm else ""
        powerstate = str(row[norm.get("Powerstate", 0)] or "") if "Powerstate" in norm else ""

        os_type = detect_os_type(os_config, os_tools)
        desktop = is_desktop_os(os_config, os_tools)
        is_sql = detect_sql(str(vm_name), annotation, folder, cluster, os_type)
        source_file = str(row[norm.get("Source File", 0)] or "") if "Source File" in norm else ""
        is_nonprod = detect_nonprod(str(vm_name), folder, cluster, source_file)

        mem_gib = round(int(mem_mib) / 1024, 1)
        disk_gib = round(int(disk_mib) / 1024, 0) if disk_mib else 0

        os_detail = os_tools if os_tools and os_tools != "None" else os_config
        linux_distro = detect_linux_distro(os_detail) if os_type == "Linux" else ""

        vms.append({
            "vm_name": str(vm_name),
            "powerstate": powerstate,
            "source_cpus": int(cpus),
            "source_mem_gib": mem_gib,
            "source_mem_mib": int(mem_mib),
            "source_disk_gib": int(disk_gib),
            "num_disks": int(row[norm.get("Disks", 0)] or 0) if "Disks" in norm else 0,
            "os_type": os_type,
            "os_detail": os_detail,
            "linux_distro": linux_distro,
            "is_desktop": desktop,
            "is_sql": is_sql,
            "is_nonprod": is_nonprod,
            "environment": "Non Prod" if is_nonprod else "Production",
            "annotation": annotation[:80] if annotation else "",
            "folder": folder,
            "cluster": cluster,
            "dns_name": str(row[norm.get("DNS Name", 0)] or "") if "DNS Name" in norm else "",
            "ip_address": str(row[norm.get("Primary IP Address", 0)] or "") if "Primary IP Address" in norm else "",
            "efi_secure": str(row[norm.get("EFI Secure boot", 0)] or "") if "EFI Secure boot" in norm else "",
        })

    wb.close()
    return vms

# =============================================================================
# PRICING ENGINE
# =============================================================================

def price_vm(vm):
    """Calculate all pricing tiers for a single VM."""
    src_cpu = vm["source_cpus"]
    src_mem = vm["source_mem_gib"]

    # Apply downsizing
    target_cpu, target_mem = apply_downsizing(src_cpu, src_mem)
    vm["target_cpus"] = target_cpu
    vm["target_mem_gib"] = target_mem
    vm["was_downsized"] = (target_cpu != src_cpu or target_mem != src_mem)

    # Find cheapest VM (compares by total LI price including license premiums)
    azure_vm = find_cheapest_vm(target_cpu, target_mem, vm["os_type"] == "Windows", vm["is_nonprod"], vm.get("linux_distro", "free"), PREFER_HIGHER_RAM)

    if not azure_vm:
        vm["azure_sku"] = "NO MATCH"
        vm["azure_series"] = ""
        vm["azure_vcpus"] = 0
        vm["azure_mem_gib"] = 0
        for key in ["payg_li", "payg_byol", "1yr_li", "1yr_byol", "3yr_li", "3yr_byol", "3yr_ri_li", "3yr_ri_byol"]:
            vm[f"compute_{key}"] = 0
        vm["sql_monthly"] = 0
        vm["disk_tier"] = ""
        vm["disk_size_gib"] = 0
        vm["disk_monthly"] = 0
        return vm

    vm["azure_sku"] = azure_vm["sku"]
    vm["azure_series"] = azure_vm["series"]
    vm["azure_vcpus"] = azure_vm["vcpus"]
    vm["azure_mem_gib"] = azure_vm["mem_gib"]

    is_win = vm["os_type"] == "Windows"

    # Base compute costs (BYOL = Linux rate for all OS types)
    base_payg = azure_vm["linux_payg"]
    base_1yr = azure_vm["linux_1yr"]
    base_3yr = azure_vm["linux_3yr"]

    if is_win:
        # Windows LI = Windows rate from pricing DB
        li_payg = azure_vm["win_payg"]
        li_1yr = azure_vm["win_1yr"]
        li_3yr = azure_vm["win_3yr"]
    else:
        distro = vm.get("linux_distro", "free")
        if distro == "rhel":
            # RHEL LI = actual RHEL price from Azure RHEL pricing page (per-SKU)
            li_payg = azure_vm["rhel_payg"]
            li_1yr = azure_vm["rhel_1yr"]
            li_3yr = azure_vm["rhel_3yr"]
        elif distro == "suse":
            # SUSE LI = actual SUSE price from Azure SUSE pricing page (per-SKU)
            li_payg = azure_vm["suse_payg"]
            li_1yr = azure_vm["suse_1yr"]
            li_3yr = azure_vm["suse_3yr"]
        else:
            # Free Linux (Ubuntu, Debian, Oracle, etc.) = base Linux rate
            li_payg = base_payg
            li_1yr = base_1yr
            li_3yr = base_3yr

    vm["compute_payg_li"] = round(li_payg, 2)
    vm["compute_payg_byol"] = round(base_payg, 2)
    vm["compute_1yr_li"] = round(li_1yr, 2)
    vm["compute_1yr_byol"] = round(base_1yr, 2)
    vm["compute_3yr_li"] = round(li_3yr, 2)
    vm["compute_3yr_byol"] = round(base_3yr, 2)

    # Reserved Instance 3yr pricing (uses RI lookup, falls back to SP if not found)
    ri_3yr_base = RI_3YR_PRICES.get(azure_vm["sku"], base_3yr)
    if is_win:
        # RI LI for Windows: RI base + Windows premium (win_3yr - linux_3yr)
        win_premium_3yr = azure_vm["win_3yr"] - azure_vm["linux_3yr"]
        vm["compute_3yr_ri_li"] = round(ri_3yr_base + win_premium_3yr, 2)
    else:
        distro = vm.get("linux_distro", "free")
        if distro == "rhel":
            rhel_premium_3yr = azure_vm["rhel_3yr"] - azure_vm["linux_3yr"]
            vm["compute_3yr_ri_li"] = round(ri_3yr_base + rhel_premium_3yr, 2)
        elif distro == "suse":
            suse_premium_3yr = azure_vm["suse_3yr"] - azure_vm["linux_3yr"]
            vm["compute_3yr_ri_li"] = round(ri_3yr_base + suse_premium_3yr, 2)
        else:
            vm["compute_3yr_ri_li"] = round(ri_3yr_base, 2)
    vm["compute_3yr_ri_byol"] = round(ri_3yr_base, 2)

    # SQL Standard license cost ($73/vCPU, min 4 cores)
    SQL_PER_VCPU = 73.00
    if vm["is_sql"]:
        sql_vcpus = max(azure_vm["vcpus"], 4)
        vm["sql_monthly"] = round(sql_vcpus * SQL_PER_VCPU, 2)
    else:
        vm["sql_monthly"] = 0.0

    # Storage
    disk = find_disk_tier(vm["source_disk_gib"])
    vm["disk_tier"] = disk["tier"]
    vm["disk_size_gib"] = disk["size_gib"]
    vm["disk_monthly"] = disk["monthly"]

    return vm

# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def generate_excel(vms, output_path, desktop_vms=None):
    """Generate the formatted Excel output workbook."""
    if desktop_vms is None:
        desktop_vms = []
    wb = Workbook()

    # ---- Styles ----
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    subheader_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    data_font = Font(name="Arial", size=10)
    money_fmt = '#,##0.00'
    int_fmt = '#,##0'
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    alt_fill = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")
    sql_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    nonprod_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    downsized_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    # =========================================================================
    # SUMMARY SHEET
    # =========================================================================
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.sheet_properties.tabColor = "1F3864"

    # Title
    ws_sum["A1"] = "Azure Pricing Estimate Summary"
    ws_sum["A1"].font = Font(name="Arial", size=16, bold=True, color="1F3864")
    ws_sum["A2"] = f"Generated: {datetime.now().strftime('%B %d, %Y %I:%M %p')}"
    ws_sum["A2"].font = Font(name="Arial", size=10, color="666666")
    mode = "Like-for-Like (no downsizing)" if LIKE_FOR_LIKE else f"N-1 Downsizing (CPU>{DOWNSIZE_CPU_THRESHOLD}, MEM>{DOWNSIZE_MEM_THRESHOLD})"
    ws_sum["A3"] = f"Region: East US | Storage: Standard SSD | Pricing: Microsoft Azure Website (March 2026) | Commitment: Savings Plan | Mode: {mode}"
    ws_sum["A3"].font = Font(name="Arial", size=10, color="666666")
    ws_sum["A4"] = "LI columns include RHEL/SUSE license premiums where applicable. BYOL columns are base Linux rate."
    ws_sum["A4"].font = Font(name="Arial", size=10, color="666666")

    # Summary stats
    total_vms = len(vms)
    win_vms = sum(1 for v in vms if v["os_type"] == "Windows")
    linux_vms = total_vms - win_vms
    sql_vms = sum(1 for v in vms if v["is_sql"])
    prod_vms = sum(1 for v in vms if not v["is_nonprod"])
    nonprod_vms = total_vms - prod_vms
    downsized_vms = sum(1 for v in vms if v["was_downsized"])
    rhel_vms = sum(1 for v in vms if v.get("linux_distro") == "rhel")
    suse_vms = sum(1 for v in vms if v.get("linux_distro") == "suse")
    oracle_vms = sum(1 for v in vms if v.get("linux_distro") == "oracle")
    free_linux = sum(1 for v in vms if v.get("linux_distro") == "free")

    stats = [
        ("Total VMs (Compute)", total_vms),
        ("Desktop OS (AVD — see separate tab)", len(desktop_vms)),
        ("Windows Server VMs", win_vms),
        ("Linux VMs", linux_vms),
        ("  RHEL / CentOS", rhel_vms),
        ("  SUSE / SLES", suse_vms),
        ("  Oracle Linux", oracle_vms),
        ("  Free Linux (Ubuntu, Debian, etc.)", free_linux),
        ("SQL Workloads (Detected)", sql_vms),
        ("Production VMs", prod_vms),
        ("Non Production VMs", nonprod_vms),
        ("Downsized VMs", downsized_vms),
    ]

    row = 6
    for label, val in stats:
        ws_sum[f"A{row}"] = label
        ws_sum[f"A{row}"].font = Font(name="Arial", size=10, bold=not label.startswith("  "))
        ws_sum[f"B{row}"] = val
        ws_sum[f"B{row}"].font = data_font
        ws_sum[f"B{row}"].number_format = int_fmt
        row += 1

    # Cost summary table
    row += 1
    cost_header_row = row
    cost_headers = ["Cost Category", "Pay As You Go (LI)", "Pay As You Go (BYOL)",
                    "1 Yr SP (LI)", "1 Yr SP (BYOL)", "3 Yr SP (LI)", "3 Yr SP (BYOL)",
                    "3 Yr RI (LI)", "3 Yr RI + Hybrid",
                    "SQL Std License"]
    for c, h in enumerate(cost_headers, 1):
        cell = ws_sum.cell(row=row, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    # Monthly compute
    row += 1
    ws_sum.cell(row=row, column=1, value="Monthly Compute").font = Font(name="Arial", size=10, bold=True)
    for c, key in enumerate(["payg_li", "payg_byol", "1yr_li", "1yr_byol", "3yr_li", "3yr_byol", "3yr_ri_li", "3yr_ri_byol"], 2):
        total = sum(v[f"compute_{key}"] for v in vms)
        cell = ws_sum.cell(row=row, column=c, value=total)
        cell.number_format = '$#,##0.00'
        cell.font = data_font
    # SQL column
    total_sql = sum(v["sql_monthly"] for v in vms)
    cell = ws_sum.cell(row=row, column=10, value=total_sql)
    cell.number_format = '$#,##0.00'
    cell.font = data_font

    # Monthly storage
    row += 1
    ws_sum.cell(row=row, column=1, value="Monthly Storage (Standard SSD)").font = Font(name="Arial", size=10, bold=True)
    total_storage = sum(v["disk_monthly"] for v in vms)
    for c in range(2, 10):
        cell = ws_sum.cell(row=row, column=c, value=total_storage)
        cell.number_format = '$#,##0.00'
        cell.font = data_font
    # SQL doesn't affect storage
    ws_sum.cell(row=row, column=10, value=0).number_format = '$#,##0.00'

    # Monthly total
    row += 1
    ws_sum.cell(row=row, column=1, value="Monthly Total").font = Font(name="Arial", size=11, bold=True, color="1F3864")
    for c, key in enumerate(["payg_li", "payg_byol", "1yr_li", "1yr_byol", "3yr_li", "3yr_byol", "3yr_ri_li", "3yr_ri_byol"], 2):
        total = sum(v[f"compute_{key}"] for v in vms) + total_storage
        cell = ws_sum.cell(row=row, column=c, value=total)
        cell.number_format = '$#,##0.00'
        cell.font = Font(name="Arial", size=11, bold=True, color="1F3864")
        cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    # SQL total
    cell = ws_sum.cell(row=row, column=10, value=total_sql)
    cell.number_format = '$#,##0.00'
    cell.font = Font(name="Arial", size=11, bold=True, color="1F3864")
    cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")

    # Annual total
    row += 1
    ws_sum.cell(row=row, column=1, value="Annual Total").font = Font(name="Arial", size=11, bold=True, color="C00000")
    for c, key in enumerate(["payg_li", "payg_byol", "1yr_li", "1yr_byol", "3yr_li", "3yr_byol", "3yr_ri_li", "3yr_ri_byol"], 2):
        total = (sum(v[f"compute_{key}"] for v in vms) + total_storage) * 12
        cell = ws_sum.cell(row=row, column=c, value=total)
        cell.number_format = '$#,##0.00'
        cell.font = Font(name="Arial", size=11, bold=True, color="C00000")
        cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    # SQL annual
    cell = ws_sum.cell(row=row, column=10, value=total_sql * 12)
    cell.number_format = '$#,##0.00'
    cell.font = Font(name="Arial", size=11, bold=True, color="C00000")
    cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    # Column widths for summary
    ws_sum.column_dimensions["A"].width = 42
    for col_letter in ["B", "C", "D", "E", "F", "G", "H", "I", "J"]:
        ws_sum.column_dimensions[col_letter].width = 22

    # =========================================================================
    # DETAIL SHEET
    # =========================================================================
    ws_det = wb.create_sheet("VM Detail")
    ws_det.sheet_properties.tabColor = "2E75B6"

    detail_headers = [
        "VM Name",                          # A
        "Environment",                      # B
        "OS Type",                          # C
        "OS Detail",                        # D
        "Linux Distro",                     # E
        "SQL Workload",                     # F
        "Source vCPUs",                     # G
        "Source Memory (GiB)",              # H
        "Source Disk (GiB)",                # I
        "Target vCPUs (After Downsizing)",  # J
        "Target Memory (GiB)",              # K
        "Downsized",                        # L
        "Azure SKU",                        # M
        "Azure Series",                     # N
        "Azure vCPUs",                      # O
        "Azure Memory (GiB)",              # P
        "Compute PAYG LI (Monthly)",        # Q
        "Compute PAYG BYOL (Monthly)",      # R
        "Compute 1Yr SP LI (Monthly)",      # S
        "Compute 1Yr SP BYOL (Monthly)",    # T
        "Compute 3Yr SP LI (Monthly)",      # U
        "Compute 3Yr SP BYOL (Monthly)",    # V
        "Compute 3Yr RI LI (Monthly)",      # W
        "Compute 3Yr RI+Hybrid (Monthly)",  # X
        "SQL Std License (Monthly)",        # Y
        "Disk Tier",                        # Z
        "Disk Size (GiB)",                  # AA
        "Disk Monthly Cost",                # AB
        "Total Monthly (PAYG LI)",          # AC
        "Total Monthly (3Yr SP BYOL)",      # AD
        "Total Monthly (3Yr RI+Hybrid)",    # AE
        "Folder",                           # AF
        "Cluster",                          # AG
    ]

    for c, h in enumerate(detail_headers, 1):
        cell = ws_det.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin_border

    for r, vm in enumerate(vms, 2):
        is_alt = r % 2 == 0
        row_fill = alt_fill if is_alt else None

        values = [
            vm["vm_name"],
            vm["environment"],
            vm["os_type"],
            vm["os_detail"][:50],
            vm.get("linux_distro", ""),
            "Yes" if vm["is_sql"] else "No",
            vm["source_cpus"],
            vm["source_mem_gib"],
            vm["source_disk_gib"],
            vm["target_cpus"],
            vm["target_mem_gib"],
            "Yes" if vm["was_downsized"] else "No",
            vm["azure_sku"],
            vm["azure_series"],
            vm["azure_vcpus"],
            vm["azure_mem_gib"],
            vm["compute_payg_li"],
            vm["compute_payg_byol"],
            vm["compute_1yr_li"],
            vm["compute_1yr_byol"],
            vm["compute_3yr_li"],
            vm["compute_3yr_byol"],
            vm["compute_3yr_ri_li"],
            vm["compute_3yr_ri_byol"],
            vm["sql_monthly"],
            vm["disk_tier"],
            vm["disk_size_gib"],
            vm["disk_monthly"],
            vm["compute_payg_li"] + vm["disk_monthly"],
            vm["compute_3yr_byol"] + vm["disk_monthly"],
            vm["compute_3yr_ri_byol"] + vm["disk_monthly"],
            vm["folder"],
            vm["cluster"],
        ]

        for c, val in enumerate(values, 1):
            cell = ws_det.cell(row=r, column=c, value=val)
            cell.font = data_font
            cell.border = thin_border

            if isinstance(val, float):
                cell.number_format = money_fmt
            elif isinstance(val, int) and c >= 7:
                cell.number_format = int_fmt

            # Conditional fills
            if vm["is_sql"] and c == 6:
                cell.fill = sql_fill
            elif vm["is_nonprod"] and c == 2:
                cell.fill = nonprod_fill
            elif vm["was_downsized"] and c == 12:
                cell.fill = downsized_fill
            elif row_fill:
                cell.fill = row_fill

    # Column widths for detail
    col_widths = {
        "A": 28, "B": 14, "C": 10, "D": 32, "E": 12, "F": 12,
        "G": 12, "H": 16, "I": 14, "J": 22, "K": 18,
        "L": 12, "M": 24, "N": 12, "O": 12, "P": 16,
        "Q": 20, "R": 22, "S": 22, "T": 22, "U": 22, "V": 22,
        "W": 24, "X": 26, "Y": 22,
        "Z": 10, "AA": 14, "AB": 16, "AC": 22, "AD": 24, "AE": 26,
        "AF": 40, "AG": 20,
    }
    for col, width in col_widths.items():
        ws_det.column_dimensions[col].width = width

    # Freeze header
    ws_det.freeze_panes = "A2"

    # Auto filter
    ws_det.auto_filter.ref = f"A1:{get_column_letter(len(detail_headers))}{len(vms) + 1}"

    # =========================================================================
    # SQL WORKLOADS SHEET
    # =========================================================================
    sql_list = [v for v in vms if v["is_sql"]]
    if sql_list:
        ws_sql = wb.create_sheet("SQL Workloads")
        ws_sql.sheet_properties.tabColor = "FFC000"

        sql_headers = ["VM Name", "Environment", "OS Detail", "Source vCPUs", "Source Memory (GiB)",
                       "Source Disk (GiB)", "Azure SKU", "Azure vCPUs",
                       "PAYG LI (Monthly)", "PAYG BYOL (Monthly)",
                       "1Yr SP LI (Monthly)", "1Yr SP BYOL (Monthly)", "3Yr SP LI (Monthly)",
                       "3Yr SP BYOL (Monthly)", "3Yr RI LI (Monthly)", "3Yr RI+Hybrid (Monthly)",
                       "SQL Std License (Monthly)",
                       "Disk Monthly", "Total Monthly (3Yr RI+Hybrid + SQL)"]

        for c, h in enumerate(sql_headers, 1):
            cell = ws_sql.cell(row=1, column=c, value=h)
            cell.font = header_font
            cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
            cell.font = Font(name="Arial", size=10, bold=True, color="000000")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        for r, vm in enumerate(sql_list, 2):
            vals = [vm["vm_name"], vm["environment"], vm["os_detail"][:50],
                    vm["source_cpus"], vm["source_mem_gib"], vm["source_disk_gib"],
                    vm["azure_sku"], vm["azure_vcpus"],
                    vm["compute_payg_li"], vm["compute_payg_byol"],
                    vm["compute_1yr_li"], vm["compute_1yr_byol"],
                    vm["compute_3yr_li"], vm["compute_3yr_byol"],
                    vm["compute_3yr_ri_li"], vm["compute_3yr_ri_byol"],
                    vm["sql_monthly"],
                    vm["disk_monthly"],
                    vm["compute_3yr_ri_byol"] + vm["sql_monthly"] + vm["disk_monthly"]]
            for c, val in enumerate(vals, 1):
                cell = ws_sql.cell(row=r, column=c, value=val)
                cell.font = data_font
                if isinstance(val, float):
                    cell.number_format = money_fmt

    # =========================================================================
    # PRICING REFERENCE SHEET
    # =========================================================================
    ws_ref = wb.create_sheet("Pricing Reference")
    ws_ref.sheet_properties.tabColor = "70AD47"

    ws_ref["A1"] = "Azure VM Pricing Reference (East US)"
    ws_ref["A1"].font = Font(name="Arial", size=14, bold=True, color="1F3864")
    ws_ref["A2"] = "This sheet shows the pricing data used by the tool. Update these values from the Azure website for latest pricing."
    ws_ref["A2"].font = Font(name="Arial", size=10, color="666666")

    ref_headers = ["SKU", "Series", "vCPUs", "Memory (GiB)",
                   "Linux PAYG/mo", "Windows PAYG/mo",
                   "Linux 1Yr SP/mo", "Windows 1Yr SP/mo",
                   "Linux 3Yr SP/mo", "Windows 3Yr SP/mo"]

    for c, h in enumerate(ref_headers, 1):
        cell = ws_ref.cell(row=4, column=c, value=h)
        cell.font = header_font
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for r, vm in enumerate(AZURE_VM_PRICING, 5):
        vals = [vm["sku"], vm["series"], vm["vcpus"], vm["mem_gib"],
                vm["linux_payg"], vm["win_payg"],
                vm["linux_1yr"], vm["win_1yr"],
                vm["linux_3yr"], vm["win_3yr"]]
        for c, val in enumerate(vals, 1):
            cell = ws_ref.cell(row=r, column=c, value=val)
            cell.font = data_font
            if isinstance(val, float):
                cell.number_format = '$#,##0.00'

    for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]:
        ws_ref.column_dimensions[col].width = 20

    # =========================================================================
    # DISK PRICING REFERENCE
    # =========================================================================
    disk_start_row = 5 + len(AZURE_VM_PRICING) + 3
    ws_ref.cell(row=disk_start_row, column=1, value="Standard SSD Managed Disk Pricing (East US)").font = Font(name="Arial", size=12, bold=True, color="1F3864")

    disk_headers = ["Disk Tier", "Size (GiB)", "Monthly Cost"]
    for c, h in enumerate(disk_headers, 1):
        cell = ws_ref.cell(row=disk_start_row + 1, column=c, value=h)
        cell.font = header_font
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    for r, d in enumerate(MANAGED_DISK_PRICING, disk_start_row + 2):
        ws_ref.cell(row=r, column=1, value=d["tier"]).font = data_font
        ws_ref.cell(row=r, column=2, value=d["size_gib"]).font = data_font
        cell = ws_ref.cell(row=r, column=3, value=d["monthly"])
        cell.font = data_font
        cell.number_format = '$#,##0.00'

    # =========================================================================
    # RHEL/SUSE PREMIUM REFERENCE
    # =========================================================================
    prem_start_row = disk_start_row + len(MANAGED_DISK_PRICING) + 4
    ws_ref.cell(row=prem_start_row, column=1, value="RHEL License Premium (Monthly, added to LI columns)").font = Font(name="Arial", size=12, bold=True, color="1F3864")
    prem_headers = ["vCPUs", "PAYG Premium", "1Yr SP Premium", "3Yr SP Premium"]
    for c, h in enumerate(prem_headers, 1):
        cell = ws_ref.cell(row=prem_start_row + 1, column=c, value=h)
        cell.font = header_font
        cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    for i, vcpu in enumerate(sorted(RHEL_PREMIUM.keys())):
        r = prem_start_row + 2 + i
        ws_ref.cell(row=r, column=1, value=vcpu).font = data_font
        for c, val in enumerate(RHEL_PREMIUM[vcpu], 2):
            cell = ws_ref.cell(row=r, column=c, value=val)
            cell.font = data_font
            cell.number_format = '$#,##0.00'

    suse_start = prem_start_row + 2 + len(RHEL_PREMIUM) + 2
    ws_ref.cell(row=suse_start, column=1, value="SUSE License Premium (Monthly, added to LI columns)").font = Font(name="Arial", size=12, bold=True, color="1F3864")
    for c, h in enumerate(prem_headers, 1):
        cell = ws_ref.cell(row=suse_start + 1, column=c, value=h)
        cell.font = header_font
        cell.fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    for i, vcpu in enumerate(sorted(SUSE_PREMIUM.keys())):
        r = suse_start + 2 + i
        ws_ref.cell(row=r, column=1, value=vcpu).font = data_font
        for c, val in enumerate(SUSE_PREMIUM[vcpu], 2):
            cell = ws_ref.cell(row=r, column=c, value=val)
            cell.font = data_font
            cell.number_format = '$#,##0.00'

    # =========================================================================
    # INSTANCE CONSOLIDATION TAB
    # =========================================================================
    ws_con = wb.create_sheet("Instance Consolidation")
    ws_con.sheet_properties.tabColor = "7030A0"

    # Group VMs by Azure SKU and count
    from collections import Counter, defaultdict
    sku_counts = Counter()
    sku_info = {}
    for vm in vms:
        sku = vm.get("azure_sku", "NO MATCH")
        if sku == "NO MATCH":
            continue
        sku_counts[sku] += 1
        if sku not in sku_info:
            # Find the SKU in pricing DB to get specs
            for p in AZURE_VM_PRICING:
                if p["sku"] == sku:
                    sku_info[sku] = {"vcpus": p["vcpus"], "mem_gib": p["mem_gib"]}
                    break

    # Title
    ws_con["A1"] = "Instance Type Consolidation"
    ws_con["A1"].font = Font(name="Arial", size=14, bold=True, color="1F3864")
    ws_con["A2"] = f"Total VMs: {sum(sku_counts.values())} | Unique Instance Types: {len(sku_counts)}"
    ws_con["A2"].font = Font(name="Arial", size=10, color="666666")

    # Headers
    con_headers = ["VM Size", "vCPU", "RAM (GB)", "Count of VMs"]
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    for c, h in enumerate(con_headers, 1):
        cell = ws_con.cell(row=4, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Data rows — sorted by SKU name to match Sasi's format
    row = 5
    total_vms = 0
    for sku in sorted(sku_counts.keys()):
        info = sku_info.get(sku, {"vcpus": 0, "mem_gib": 0})
        # Clean SKU name: "Standard_B4as_v2" → "B4as v2"
        display_name = sku.replace("Standard_", "").replace("_", " ")
        count = sku_counts[sku]
        total_vms += count

        ws_con.cell(row=row, column=1, value=display_name).font = data_font
        ws_con.cell(row=row, column=2, value=info["vcpus"]).font = data_font
        ws_con.cell(row=row, column=3, value=info["mem_gib"]).font = data_font
        cell = ws_con.cell(row=row, column=4, value=count)
        cell.font = data_font
        cell.number_format = '#,##0'

        # Alternate row shading
        if row % 2 == 0:
            light_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            for c in range(1, 5):
                ws_con.cell(row=row, column=c).fill = light_fill
        row += 1

    # Total row
    ws_con.cell(row=row, column=1, value="TOTAL").font = Font(name="Arial", size=10, bold=True, color="1F3864")
    cell = ws_con.cell(row=row, column=4, value=total_vms)
    cell.font = Font(name="Arial", size=10, bold=True, color="1F3864")
    cell.number_format = '#,##0'

    # Column widths
    ws_con.column_dimensions["A"].width = 22
    ws_con.column_dimensions["B"].width = 10
    ws_con.column_dimensions["C"].width = 12
    ws_con.column_dimensions["D"].width = 16

    ws_con.freeze_panes = "A5"

    # =========================================================================
    # AZURE VIRTUAL DESKTOP (AVD) TAB — Desktop OS VMs
    # =========================================================================
    if desktop_vms:
        ws_avd = wb.create_sheet("Azure Virtual Desktop")
        ws_avd.sheet_properties.tabColor = "00B0F0"

        ws_avd["A1"] = "Azure Virtual Desktop (AVD) Candidates"
        ws_avd["A1"].font = Font(name="Arial", size=14, bold=True, color="1F3864")
        ws_avd["A2"] = f"Desktop OS VMs excluded from compute pricing — to be priced as AVD Workspaces separately"
        ws_avd["A2"].font = Font(name="Arial", size=10, color="666666")
        ws_avd["A3"] = f"Total Desktop VMs: {len(desktop_vms)}"
        ws_avd["A3"].font = Font(name="Arial", size=10, bold=True, color="C00000")

        avd_headers = ["VM Name", "OS", "vCPUs", "Memory (GiB)", "Disk (GiB)", "Environment", "Source vCenter"]
        for c, h in enumerate(avd_headers, 1):
            cell = ws_avd.cell(row=5, column=c, value=h)
            cell.font = header_font
            cell.fill = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")

        for i, vm in enumerate(desktop_vms):
            r = 6 + i
            vals = [
                vm["vm_name"],
                vm.get("os_detail", ""),
                vm["source_cpus"],
                vm["source_mem_gib"],
                vm["source_disk_gib"],
                vm.get("environment", ""),
                "",
            ]
            for c, val in enumerate(vals, 1):
                cell = ws_avd.cell(row=r, column=c, value=val)
                cell.font = data_font
                if i % 2 == 1:
                    cell.fill = PatternFill(start_color="E8F4FD", end_color="E8F4FD", fill_type="solid")

        ws_avd.column_dimensions["A"].width = 55
        ws_avd.column_dimensions["B"].width = 40
        ws_avd.column_dimensions["C"].width = 10
        ws_avd.column_dimensions["D"].width = 14
        ws_avd.column_dimensions["E"].width = 12
        ws_avd.column_dimensions["F"].width = 14
        ws_avd.column_dimensions["G"].width = 22
        ws_avd.freeze_panes = "A6"

    wb.save(output_path)
    return output_path

# =============================================================================
# AZURE MIGRATE CSV EXPORT
# =============================================================================


# =============================================================================
# MAIN
# =============================================================================

def parse_adhoc_file(filepath):
    """Parse simplified VM list format (Book1/optimised style).
    Expected columns: VM, Prod, CPU, Memory, OS (or similar), SQL
    Auto-detects column layout."""
    wb = load_workbook(filepath, read_only=True, data_only=True)
    #ws = wb[wb.sheetnames[0]]#
    PREFERRED = ["VM List", "VM list", "Primary (Inscope)", "In-Scope", "Sheet1"]
    sheet_name = next((s for s in PREFERRED if s in wb.sheetnames), wb.sheetnames[0])
    ws = wb[sheet_name]
    print(f"  Reading sheet: '{sheet_name}'")
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    header = [str(h or "").strip().lower() for h in rows[0]]

    # Auto-detect columns by fuzzy matching
    col = {}
    for i, h in enumerate(header):
        if h in ("vm", "vm name", "name"): col["vm"] = i
        elif "prod" in h: col["prod"] = i
        elif "cpu" in h: col["cpu"] = i
        elif "mem" in h: col["mem"] = i
        elif "os" in h: col["os"] = i
        elif "sql" in h: col["sql"] = i
        elif "cluster" in h: col["cluster"] = i
        elif "datacenter" in h: col["datacenter"] = i

    # Handle two known formats:
    # Book1: VM, CPU, Memory, OS, Prod, Sql (6 cols)
    # Optimised: VM, Prod, CPU to be considered, memory to be considered, Datacenter, Cluster, OS according to VMware Tools, SQL (8 cols)
    if "vm" not in col:
        col["vm"] = 0
    if "cpu" not in col:
        # Try positional fallback for Book1 format
        col["cpu"] = 1
        col["mem"] = 2
        col["os"] = 3
        col["prod"] = 4
        col["sql"] = 5

    import math
    vms = []
    for r in rows[1:]:
        if not r[col["vm"]]: continue
        vm_name = str(r[col["vm"]])
        cpu = int(float(r[col.get("cpu", 1)] or 0))
        mem_raw = float(r[col.get("mem", 2)] or 0)
        mem = math.ceil(mem_raw)

        os_str = str(r[col.get("os", 3)] or "")
        prod_val = str(r[col.get("prod", 4)] or "").strip().lower()
        sql_val = str(r[col.get("sql", 5)] or "").strip().upper()

        # Detect OS
        ol = os_str.lower()
        if "windows" in ol:
            os_type = "Windows"; distro = ""
        elif "red hat" in ol or "rhel" in ol or "centos" in ol:
            os_type = "Linux"; distro = "rhel"
        elif "suse" in ol or "sles" in ol:
            os_type = "Linux"; distro = "suse"
        elif "oracle" in ol:
            os_type = "Linux"; distro = "oracle"
        else:
            os_type = "Linux"; distro = "free"

        is_prod = prod_val in ("yes", "true", "1", "prod", "production")
        has_sql = sql_val in ("SQL", "SQL STD", "SQL STANDARD", "YES")

        vms.append({
            "vm_name": vm_name,
            "source_cpus": cpu,
            "source_mem_gib": mem_raw,
            "target_cpus": cpu,
            "target_mem_gib": mem,
            "source_disk_gib": 0,
            "os_type": os_type,
            "os_detail": os_str,
            "linux_distro": distro,
            "is_nonprod": not is_prod,
            "is_sql": has_sql,
            "environment": "Production" if is_prod else "Non-Prod",
            "powerstate": "",
            "num_disks": 0,
        })

    return vms


def generate_adhoc_excel(vms, output_path, pricing_mode="li", include_sql=True, use_ri=False):
    """Generate Excel for adhoc files. pricing_mode: 'li' or 'byol'. use_ri: label as 3yr RI + Hybrid."""
    from datetime import datetime
    wb = Workbook()

    SQL_PER_VCPU = 73.00  # 3yr SP SQL Standard per vCPU (same across PAYG/1yr/3yr)

    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    data_font = Font(name="Arial", size=10)
    money_fmt = '$#,##0.00'
    sql_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Determine labels based on mode
    if use_ri:
        pricing_title = "3yr Reserved Instance + Azure Hybrid Benefit"
        col_label_3yr = "3Yr RI + Hybrid"
        col_label_payg = "PAYG BYOL"
        col_label_1yr = "1Yr RI + Hybrid"
    elif pricing_mode == "li":
        pricing_title = "License Included"
        col_label_3yr = "3Yr SP LI"
        col_label_payg = "PAYG LI"
        col_label_1yr = "1Yr SP LI"
    else:
        pricing_title = "BYOL (Savings Plan)"
        col_label_3yr = "3Yr SP BYOL"
        col_label_payg = "PAYG BYOL"
        col_label_1yr = "1Yr SP BYOL"

    # Price each VM
    for vm in vms:
        candidates = []
        for az in AZURE_VM_PRICING:
            if az["vcpus"] >= vm["target_cpus"] and az["mem_gib"] >= vm["target_mem_gib"]:
                if not vm["is_nonprod"] and az["series"] == "B":
                    continue
                candidates.append(az)

        if not candidates:
            vm["azure_sku"] = "NO MATCH"; vm["azure_series"] = ""; vm["azure_vcpus"] = 0; vm["azure_mem"] = 0
            vm["payg"] = vm["yr1"] = vm["yr3"] = 0
        else:
            if pricing_mode == "li":
                is_win = vm["os_type"] == "Windows"
                key = "win" if is_win else "linux"
                candidates.sort(key=lambda x: x[f"{key}_payg"])
                best = candidates[0]
                vm["payg"] = best[f"{key}_payg"]
                vm["yr1"] = best[f"{key}_1yr"]
                vm["yr3"] = best[f"{key}_3yr"]
            else:
                # For RI mode, sort by RI 3yr price to pick cheapest at RI rates
                if use_ri:
                    candidates.sort(key=lambda x: RI_3YR_PRICES.get(x["sku"], x["linux_3yr"]))
                else:
                    candidates.sort(key=lambda x: x["linux_payg"])
                best = candidates[0]
                vm["payg"] = best["linux_payg"]
                vm["yr1"] = best["linux_1yr"]
                if use_ri:
                    vm["yr3"] = RI_3YR_PRICES.get(best["sku"], best["linux_3yr"])
                else:
                    vm["yr3"] = best["linux_3yr"]

            vm["azure_sku"] = best["sku"].replace("Standard_", "").replace("_", " ")
            vm["azure_series"] = best["series"]
            vm["azure_vcpus"] = best["vcpus"]
            vm["azure_mem"] = best["mem_gib"]

        # SQL costs
        vm["sql_payg"] = vm["sql_1yr"] = vm["sql_3yr"] = 0
        if include_sql and vm["is_sql"] and vm.get("azure_vcpus", 0) > 0:
            sql_vcpus = max(vm["azure_vcpus"], 4)
            vm["sql_payg"] = sql_vcpus * SQL_PER_VCPU
            vm["sql_1yr"] = sql_vcpus * SQL_PER_VCPU
            vm["sql_3yr"] = sql_vcpus * SQL_PER_VCPU

        vm["total_payg"] = vm["payg"] + vm["sql_payg"]
        vm["total_1yr"] = vm["yr1"] + vm["sql_1yr"]
        vm["total_3yr"] = vm["yr3"] + vm["sql_3yr"]

    mode_label = pricing_title
    sql_label = " + SQL Standard" if include_sql and any(v["is_sql"] for v in vms) else ""

    # === Summary ===
    ws = wb.active; ws.title = "Summary"; ws.sheet_properties.tabColor = "1F3864"
    ws["A1"] = f"Azure Pricing — Like-for-Like, {mode_label}{sql_label}"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color="1F3864")
    subtitle = f"Generated: {datetime.now().strftime('%B %d, %Y %I:%M %p')} | B-series: Non-Prod Only | {mode_label}"
    if use_ri:
        subtitle += " | Note: 3yr RI rates = 3yr SP rates for AMD series (Bas v2, Das v5, Eas v5)"
    ws["A2"] = subtitle
    ws["A2"].font = Font(name="Arial", size=10, color="666666")

    stats = [
        ("Total VMs", len(vms)),
        ("Production", sum(1 for v in vms if not v["is_nonprod"])),
        ("Non-Prod", sum(1 for v in vms if v["is_nonprod"])),
        ("Windows", sum(1 for v in vms if v["os_type"] == "Windows")),
        ("RHEL", sum(1 for v in vms if v["linux_distro"] == "rhel")),
        ("SUSE", sum(1 for v in vms if v["linux_distro"] == "suse")),
        ("SQL Standard", sum(1 for v in vms if v["is_sql"])),
        ("No Match", sum(1 for v in vms if v["azure_sku"] == "NO MATCH")),
    ]
    for i, (label, val) in enumerate(stats):
        ws.cell(row=4+i, column=1, value=label).font = Font(name="Arial", size=10, bold=True)
        ws.cell(row=4+i, column=2, value=val).font = data_font

    t_payg = sum(v["total_payg"] for v in vms)
    t_1yr = sum(v["total_1yr"] for v in vms)
    t_3yr = sum(v["total_3yr"] for v in vms)

    r = 14
    for c, h in enumerate(["", "Monthly", "Annual (ARR)"], 1):
        cell = ws.cell(row=r, column=c, value=h); cell.font = header_font; cell.fill = header_fill
    for i, (label, val) in enumerate([
        (col_label_payg, t_payg), (col_label_1yr, t_1yr), (col_label_3yr, t_3yr)]):
        ws.cell(row=r+1+i, column=1, value=label).font = Font(name="Arial", size=10, bold=True)
        cell = ws.cell(row=r+1+i, column=2, value=round(val, 2)); cell.font = data_font; cell.number_format = money_fmt
        cell = ws.cell(row=r+1+i, column=3, value=round(val*12, 2)); cell.font = data_font; cell.number_format = money_fmt
    ws.column_dimensions["A"].width = 25; ws.column_dimensions["B"].width = 18; ws.column_dimensions["C"].width = 18

    # === VM Detail ===
    ws2 = wb.create_sheet("VM Detail"); ws2.sheet_properties.tabColor = "2E75B6"
    det_headers = ["VM Name", "Environment", "OS Type", "OS Detail", "Distro", "SQL",
        "Source vCPUs", "Source Memory (GiB)", "Azure SKU", "Azure vCPUs", "Azure Memory (GiB)",
        f"Compute {col_label_payg}", f"Compute {col_label_1yr}", f"Compute {col_label_3yr}",
        "SQL Std PAYG", "SQL Std 1Yr", "SQL Std 3Yr",
        f"Total {col_label_payg}", f"Total {col_label_1yr}", f"Total {col_label_3yr}"]
    for c, h in enumerate(det_headers, 1):
        cell = ws2.cell(row=1, column=c, value=h); cell.font = header_font; cell.fill = header_fill
    for i, vm in enumerate(vms):
        row = 2 + i
        vals = [vm["vm_name"], vm["environment"], vm["os_type"], vm["os_detail"], vm["linux_distro"],
            "SQL Std" if vm["is_sql"] else "NA",
            vm["source_cpus"], vm["source_mem_gib"], vm["azure_sku"], vm["azure_vcpus"], vm["azure_mem"],
            vm["payg"], vm["yr1"], vm["yr3"],
            vm["sql_payg"], vm["sql_1yr"], vm["sql_3yr"],
            vm["total_payg"], vm["total_1yr"], vm["total_3yr"]]
        for c, val in enumerate(vals, 1):
            cell = ws2.cell(row=row, column=c, value=val); cell.font = data_font
            if c >= 12: cell.number_format = money_fmt
        if vm["is_sql"]:
            for c in range(1, len(vals)+1):
                ws2.cell(row=row, column=c).fill = sql_fill
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:T{len(vms)+1}"
    for c, w in enumerate([40,12,10,42,8,8,12,16,22,12,16,14,14,14,14,14,14,14,14,14], 1):
        ws2.column_dimensions[get_column_letter(c)].width = w

    # === Instance Consolidation ===
    ws3 = wb.create_sheet("Instance Consolidation"); ws3.sheet_properties.tabColor = "7030A0"
    ws3["A1"] = "Instance Type Consolidation"; ws3["A1"].font = Font(name="Arial", size=14, bold=True, color="1F3864")
    con_headers = ["VM Size", "vCPU", "RAM (GB)", "Count of VMs"]
    for c, h in enumerate(con_headers, 1):
        cell = ws3.cell(row=3, column=c, value=h); cell.font = header_font; cell.fill = header_fill
    from collections import Counter
    sku_counts = Counter(v["azure_sku"] for v in vms if v["azure_sku"] != "NO MATCH")
    sku_info = {}
    for v in vms:
        s = v["azure_sku"]
        if s not in sku_info and s != "NO MATCH":
            sku_info[s] = {"cpu": v["azure_vcpus"], "mem": v["azure_mem"]}
    row = 4; total = 0
    for sku in sorted(sku_counts.keys()):
        info = sku_info.get(sku, {"cpu":0,"mem":0}); count = sku_counts[sku]; total += count
        ws3.cell(row=row,column=1,value=sku).font=data_font
        ws3.cell(row=row,column=2,value=info["cpu"]).font=data_font
        ws3.cell(row=row,column=3,value=info["mem"]).font=data_font
        ws3.cell(row=row,column=4,value=count).font=data_font
        row += 1
    ws3.cell(row=row,column=1,value="TOTAL").font=Font(name="Arial",size=10,bold=True)
    ws3.cell(row=row,column=4,value=total).font=Font(name="Arial",size=10,bold=True)
    ws3.column_dimensions["A"].width=22; ws3.column_dimensions["B"].width=10
    ws3.column_dimensions["C"].width=12; ws3.column_dimensions["D"].width=16

    wb.save(output_path)

    # Print summary
    no_match = sum(1 for v in vms if v["azure_sku"] == "NO MATCH")
    print(f"\n  Total VMs: {len(vms)}")
    print(f"  Production: {sum(1 for v in vms if not v['is_nonprod'])}")
    print(f"  Non-Prod: {sum(1 for v in vms if v['is_nonprod'])}")
    print(f"  SQL: {sum(1 for v in vms if v['is_sql'])}")
    print(f"  No Match: {no_match}")
    if no_match:
        for v in vms:
            if v["azure_sku"] == "NO MATCH":
                print(f"    {v['vm_name']}: {v['source_cpus']}cpu / {v['source_mem_gib']}GiB")
    print(f"\n  {col_label_payg}:  ${t_payg:>12,.2f}/mo  ARR ${t_payg*12:>14,.2f}")
    print(f"  {col_label_1yr}: ${t_1yr:>12,.2f}/mo  ARR ${t_1yr*12:>14,.2f}")
    print(f"  {col_label_3yr}: ${t_3yr:>12,.2f}/mo  ARR ${t_3yr*12:>14,.2f}")


def main():
    global LIKE_FOR_LIKE, DOWNSIZE_CPU_THRESHOLD, DOWNSIZE_MEM_THRESHOLD, PREFER_HIGHER_RAM, B_SERIES_ALL

    if len(sys.argv) < 2:
        print("Usage:")
        print("  RVTools mode:  python azure_pricing_tool.py <RVTools.xlsx> [--like-for-like] [--downsize CPU MEM] [--prefer-higher-ram] [--b-nonprod-only]")
        print("  Adhoc mode:    python azure_pricing_tool.py <file.xlsx> --adhoc [--li|--byol] [--sql] [--b-nonprod-only]")
        print("")
        print("Examples:")
        print("  python azure_pricing_tool.py Consolidated_RVTools.xlsx --like-for-like --b-nonprod-only")
        print("  python azure_pricing_tool.py Book1.xlsx --adhoc --li --sql")
        print("  python azure_pricing_tool.py optimised_file.xlsx --adhoc --byol")
        sys.exit(1)

    input_file = sys.argv[1]

    if not os.path.exists(input_file):
        print(f"Error: File not found: {input_file}")
        sys.exit(1)

    # Parse CLI flags
    args = sys.argv[2:]

    # ===== ADHOC MODE =====
    if "--adhoc" in args:
        pricing_mode = "byol"
        if "--li" in args:
            pricing_mode = "li"
        include_sql = "--sql" in args
        use_ri = "--ri" in args
        if "--b-nonprod-only" in args:
            B_SERIES_ALL = False

        # When --ri is used, force BYOL (Hybrid Benefit = Linux base rate = RI rate for AMD series)
        if use_ri:
            pricing_mode = "byol"

        mode_label = "LI" if pricing_mode == "li" else ("3yr_RI_Hybrid" if use_ri else "BYOL")
        sql_label = "_SQL" if include_sql else ""
        input_dir = os.path.dirname(os.path.abspath(input_file))
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        b_suffix = "_B_NonProdOnly" if not B_SERIES_ALL else "_B_All"
        output_file = os.path.join(input_dir, f"Azure_Pricing_{base_name}_L4L_{mode_label}{sql_label}{b_suffix}.xlsx")

        ri_display = "3yr Reserved Instance + Azure Hybrid Benefit" if use_ri else ("License Included" if pricing_mode == "li" else "BYOL (Savings Plan)")

        print("=" * 60)
        print(f"Azure Pricing Automation Tool — Adhoc Mode")
        print(f"  File: {input_file}")
        print(f"  Pricing: {ri_display} | SQL: {'Yes' if include_sql else 'No'} | B-series: {'All' if B_SERIES_ALL else 'Non-Prod Only'}")
        if use_ri:
            print(f"  Note: 3yr RI + Hybrid = Linux base rate (BYOL). RI and SP 3yr prices are identical for AMD series.")
        print("=" * 60)

        print(f"\nReading file: {input_file}")
        vms = parse_adhoc_file(input_file)

        print(f"\nGenerating Excel output: {output_file}")
        generate_adhoc_excel(vms, output_file, pricing_mode=pricing_mode, include_sql=include_sql, use_ri=use_ri)

        print(f"\nDone! Output: {output_file}")
        return output_file

    # ===== RVTOOLS MODE =====
    if "--like-for-like" in args:
        LIKE_FOR_LIKE = True
    elif "--downsize" in args:
        idx = args.index("--downsize")
        if idx + 2 < len(args):
            DOWNSIZE_CPU_THRESHOLD = int(args[idx + 1])
            DOWNSIZE_MEM_THRESHOLD = int(args[idx + 2])

    if "--prefer-higher-ram" in args:
        PREFER_HIGHER_RAM = True

    if "--b-nonprod-only" in args:
        B_SERIES_ALL = False

    # Output file naming
    input_dir = os.path.dirname(os.path.abspath(input_file))
    ram_suffix = "_HigherRAM" if PREFER_HIGHER_RAM else ""
    b_suffix = "_B_NonProdOnly" if not B_SERIES_ALL else "_B_All"
    if LIKE_FOR_LIKE:
        output_file = os.path.join(input_dir, f"Azure_Pricing_LikeForLike{b_suffix}{ram_suffix}.xlsx")
        tag = "Like-for-Like"
    else:
        output_file = os.path.join(input_dir, f"Azure_Pricing_N1_CPU{DOWNSIZE_CPU_THRESHOLD}_MEM{DOWNSIZE_MEM_THRESHOLD}{b_suffix}{ram_suffix}.xlsx")
        tag = f"N-1 (CPU>{DOWNSIZE_CPU_THRESHOLD}, MEM>{DOWNSIZE_MEM_THRESHOLD})"
    if PREFER_HIGHER_RAM:
        tag += " [Higher RAM]"
    tag += f" [B-series: {'All' if B_SERIES_ALL else 'Non-Prod Only'}]"

    print("=" * 60)
    print(f"Azure Pricing Automation Tool — {tag}")
    print("=" * 60)

    # Parse
    print(f"\nReading file: {input_file}")
    all_vms = parse_rvtools(input_file)

    # Separate desktop OS VMs (Azure Virtual Desktop candidates)
    desktop_vms = [v for v in all_vms if v.get("is_desktop", False)]
    vms = [v for v in all_vms if not v.get("is_desktop", False)]

    # Classify
    print(f"\n  Total VMs parsed:     {len(all_vms)}")
    print(f"  Desktop OS (AVD):     {len(desktop_vms)} (excluded from compute, see AVD tab)")
    print(f"  VMs to price:         {len(vms)}")
    win_count = sum(1 for v in vms if v["os_type"] == "Windows")
    linux_count = len(vms) - win_count
    sql_count = sum(1 for v in vms if v["is_sql"])
    nonprod_count = sum(1 for v in vms if v["is_nonprod"])
    rhel_count = sum(1 for v in vms if v.get("linux_distro") == "rhel")
    suse_count = sum(1 for v in vms if v.get("linux_distro") == "suse")
    oracle_count = sum(1 for v in vms if v.get("linux_distro") == "oracle")
    free_count = sum(1 for v in vms if v.get("linux_distro") == "free")
    print(f"  Windows Server:       {win_count}")
    print(f"  Linux:                {linux_count}")
    print(f"    RHEL / CentOS:      {rhel_count}")
    print(f"    SUSE / SLES:        {suse_count}")
    print(f"    Oracle Linux:       {oracle_count}")
    print(f"    Free Linux:         {free_count}")
    print(f"  SQL Detected:         {sql_count}")
    print(f"  Production:           {len(vms) - nonprod_count}")
    print(f"  Non Prod:             {nonprod_count}")

    # Price each VM
    print("\nPricing VMs...")
    for vm in vms:
        price_vm(vm)

    downsized = sum(1 for v in vms if v["was_downsized"])
    no_match = sum(1 for v in vms if v["azure_sku"] == "NO MATCH")
    print(f"  Downsized:            {downsized}")
    print(f"  No match found:       {no_match}")
    if no_match > 0:
        for v in vms:
            if v["azure_sku"] == "NO MATCH":
                print(f"    {v['vm_name']}: {v['target_cpus']} vCPU / {v['target_mem_gib']} GiB")

    # Generate output
    print(f"\nGenerating Excel output: {output_file}")
    generate_excel(vms, output_file, desktop_vms=desktop_vms)

    # Print summary
    print("\n" + "=" * 60)
    print(f"COST SUMMARY — {tag} (Monthly)")
    print("=" * 60)
    for label, key in [("Pay As You Go (LI)", "payg_li"), ("Pay As You Go (BYOL)", "payg_byol"),
                       ("1 Yr SP (LI)", "1yr_li"), ("1 Yr SP (BYOL)", "1yr_byol"),
                       ("3 Yr SP (LI)", "3yr_li"), ("3 Yr SP (BYOL)", "3yr_byol"),
                       ("3 Yr RI (LI)", "3yr_ri_li"), ("3 Yr RI + Hybrid", "3yr_ri_byol")]:
        compute = sum(v[f"compute_{key}"] for v in vms)
        storage = sum(v["disk_monthly"] for v in vms)
        total = compute + storage
        print(f"  {label:30s} Compute: ${compute:>12,.2f}  Storage: ${storage:>10,.2f}  Total: ${total:>12,.2f}")

    total_sql = sum(v["sql_monthly"] for v in vms)
    sql_count_final = sum(1 for v in vms if v["is_sql"])
    if sql_count_final > 0:
        print(f"\n  SQL Standard License ({sql_count_final} VMs):  ${total_sql:>12,.2f}/mo")

    print("\nDone!")
    return output_file

if __name__ == "__main__":
    main()

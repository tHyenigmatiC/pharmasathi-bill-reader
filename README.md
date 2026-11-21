# Bill reader

Reads photos of distributors' purchase bills **offline** and writes them into PharmaSathi's import files, so
a stack of bills is not typed by hand. Nothing is sent anywhere: the reading models run on this computer.

## Use

```bash
./run.sh /path/to/bill-photos
```

It writes into `/path/to/bill-photos/bills-read/`:

| File | What it is | What you do |
|---|---|---|
| `names.xlsx` | Every medicine name printed on the bills and the medicine it means | Check the yellow rows, correct Name / Strength / Unit / Strip size, delete the Note |
| `new-medicines.xlsx` | Medicines the shop does not have yet, in the Medicines import template | Import it in PharmaSathi (Medicines) **before** the bills |
| `bills/*.xlsx` | One purchase-template file per bill, with a **Check** column | New purchase → Import from Excel; look at yellow lines; type the overall discount the Bill details sheet gives |
| `summary.txt` | One line per photo | |

Run it again after editing `names.xlsx`: your changes are kept, and photos already read are not read again
(about a second for the whole folder). With `--medicines export.xlsx` (a Medicines export from PharmaSathi) the
shop's own medicines are matched instead of being added again.

Photos: flat, upright or sideways, in good light, the whole table in view. A bill photographed twice
(same bill number) is left out the second time.

## How it stays accurate

Reading faint print gets digits wrong (8↔3, 2↔7, 9↔4 on a worn ribbon), so the reader never trusts a single
reading. It uses the bill's own arithmetic to find and repair misreadings, and marks every value it changed or
could not check in yellow:

- each photo is read twice (small model as taken; medium model on a sharpened copy) and the reading that
  needs fewer corrections wins;
- qty × rate = amount on each line; the lines add up to the total; total − discount + rounding = net;
- the net total in words ("Six Hundred Sixty-five Only") decides the net total; rounding is never over 50 paisa;
- a discount not read is worked out from total, rounding and net;
- expiry must fall in the next few years; batch numbers must match the supplier's usual shapes (GTH1243C);
- a misread M.R.P. is repaired, or the price estimated, from the supplier's usual margin;
- a bill that needed many corrections has every line marked "compare with the paper".

One distributor's bills print Indian-rupee M.R.P.s (below the rate): the selling price is M.R.P. × 1.6. Their discount
is 3% of the amount ÷ 1.075, never the printed 3%, so the line Disc % is left blank and the bill's own
discount amount is given to type as the overall discount.

## Testing a change

`tests/truth.json` holds the true values of every test bill, read by hand and checked against each bill's
own sums. Like the photos, it is the shop's purchase data, so git leaves it out (`.gitignore`): keep a copy
with the test photos. After changing the reader:

```bash
rm -rf test-bills/bills-read/names.xlsx test-bills/bills-read/bills test-bills/bills-read/new-medicines.xlsx
./run.sh test-bills
.venv/bin/python tests/score.py tests/truth.json test-bills/bills-read
```

It prints how many values are right, wrong but marked, and **wrong and not marked**. The last must stay 0.
On 2026-11-21 (24 photos of one distributor's bills, 23 bills): 318 of 335 right (95%), 17 wrong but marked, 0 unmarked.

To add a bill to the tests: put its photo in `test-bills/` as `bill-N.jpg` and its values in `truth.json`
(batch, expiry MM/YYYY, qty, rate, amount per line; total, discount, rounding, net). Use `null` for a value
you cannot read on the paper.

## Setting it up again

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

The first run downloads the reading models (about 110 MB) into the install.

## Limits

- Only one distributor's bills have been tried. Other suppliers' column names, digit swaps and layouts may
  need adding; the usual margin and batch shapes are still learned from all suppliers together.
- Printed bills only; handwritten bills are not read.
- `test-bills/` and every `bills-read/` folder hold the shop's purchase data: keep them on this computer.

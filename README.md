# timag

Converts TIM AG quotes from xlsx into either csv or txt.

## Usage

```bash
./timag <filename.xlsx>
```

Flags:
- "--v" => switch from csv to some text format
- "--tofile" => prints not to stdout (default), but to file of same name in same folder as input, but with extension csv or txt
- "--factor 9.999" => sets an uplif factor to calculate a sales price (on the net price)
- "--po" => sets PO-column to "yes" (default: "no") to indicate this quote was finally ordered

Traverse through folders and apply conversion into central csv file:
```bash
find . -type f -name 'TIM_Angebot*.xlsx' -print0 | xargs -0 -I{} timag "{}" >>timag.csv

```

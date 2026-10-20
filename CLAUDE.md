# cbs_leyning

Builds Congregation Beth Sholom's leyning (Torah reading sign-up) sheets as
`.xlsx` workbooks from HebCal's leyning API. Usage is in `README.md`.

## Running the script

Use the `datasci` conda env, which has every dependency (requests, openpyxl,
pandas, tenacity):

```bash
~/miniforge3/envs/datasci/bin/python leyning.py --book Genesis --pages page_numbers_and_haftarot.csv
```

## Always pass the Virtual Tikkun join code when it is recorded

`vt_join_code.txt` holds Beth Sholom's Virtual Tikkun join code (one line).
**Whenever that file exists, every run that generates sheets passes it as
`--join-code`**, so each Virtual Tikkun link on the sheets ends in
`&join=<code>` and a member who signs up from a sheet joins Beth Sholom in
Virtual Tikkun instead of landing in Kahal:

```bash
~/miniforge3/envs/datasci/bin/python leyning.py --book Genesis \
  --pages page_numbers_and_haftarot.csv \
  --join-code "$(cat vt_join_code.txt)"
```

If the file is absent, run without `--join-code`. The code changes only when
a Beth Sholom admin replaces the join link in Virtual Tikkun (Admin >
Settings); then update `vt_join_code.txt` and regenerate the sheets, since
links with the old code stop joining anyone. The code is not a secret: it is
printed in every link on the published sheets.

from pathlib import Path
import datetime
import xlsxwriter


def main() -> None:
    out_dir = Path(__file__).resolve().parent
    out_dir.mkdir(parents=True, exist_ok=True)

    for old in out_dir.glob("*.xlsx"):
        old.unlink()

    build_styles_formulas(out_dir / "styles_formulas.xlsx")
    build_strings_links(out_dir / "strings_links.xlsx")
    build_multisheet_names(out_dir / "multisheet_names.xlsx")
    build_layout_breaks(out_dir / "layout_breaks.xlsx")


def build_styles_formulas(path: Path) -> None:
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Data")
    fmt_header = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
    fmt_money = wb.add_format({"num_format": "$#,##0.00", "border": 1})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})

    for col, header in enumerate(["Date", "Category", "Amount", "Tax", "Total"]):
        ws.write(0, col, header, fmt_header)

    rows = [
        ("2025-01-01", "Ops", 1200.5, 0.08),
        ("2025-01-02", "R&D", 3300.0, 0.07),
        ("2025-01-03", "Sales", 980.75, 0.09),
        ("2025-01-04", "Ops", 1450.0, 0.08),
    ]
    for row, (date_value, category, amount, tax) in enumerate(rows, start=1):
        ws.write_datetime(row, 0, datetime.datetime.fromisoformat(date_value), fmt_date)
        ws.write_string(row, 1, category)
        ws.write_number(row, 2, amount, fmt_money)
        ws.write_number(row, 3, tax)
        ws.write_formula(row, 4, f"=C{row + 1}*(1+D{row + 1})", fmt_money)

    ws.data_validation("B2:B100", {"validate": "list", "source": ["Ops", "R&D", "Sales"]})
    ws.conditional_format("E2:E100", {"type": "3_color_scale"})
    ws.merge_range(
        "G2:I2",
        "Quarterly Summary",
        wb.add_format({"align": "center", "bold": True, "bg_color": "#FCE4D6"}),
    )
    ws.write_formula("G3", "=SUM(E2:E5)", fmt_money)
    ws.write_formula("G4", "=AVERAGE(E2:E5)", fmt_money)
    ws.write_formula("G5", "=MAX(E2:E5)", fmt_money)
    ws.add_table("A1:E5", {"name": "SpendTable", "style": "Table Style Medium 2"})
    wb.close()


def build_strings_links(path: Path) -> None:
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Strings")
    for i in range(1, 200):
        ws.write(i, 0, "alpha")
        ws.write(i, 1, "beta")
        ws.write(i, 2, f"row-{i}")
    ws.write_url("D2", "https://www.sec.gov", string="SEC")
    ws.write_url("D3", "https://www.census.gov", string="Census")
    ws.write("D4", "Plain text")
    wb.close()


def build_multisheet_names(path: Path) -> None:
    wb = xlsxwriter.Workbook(str(path))
    ws_inputs = wb.add_worksheet("Inputs")
    ws_model = wb.add_worksheet("Model")
    for i in range(1, 51):
        ws_inputs.write_number(i, 0, i)
        ws_inputs.write_number(i, 1, i * 2)
    wb.define_name("InputRange", "=Inputs!$A$2:$A$51")
    ws_model.write_formula("A2", "=SUM(InputRange)")
    ws_model.write_formula("A3", "=SUM(Inputs!A2:A51)")
    ws_model.write_formula("A4", "=SUMPRODUCT(Inputs!A2:A51,Inputs!B2:B51)")
    wb.close()


def build_layout_breaks(path: Path) -> None:
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Layout")
    for row in range(120):
        for col in range(8):
            ws.write(row, col, row * col)
    ws.set_landscape()
    ws.set_paper(9)
    ws.set_margins(0.5, 0.5, 0.75, 0.75)
    ws.repeat_rows(0)
    ws.repeat_columns(0, 1)
    ws.set_h_pagebreaks([40, 80])
    ws.set_v_pagebreaks([3, 6])
    wb.close()


if __name__ == "__main__":
    main()

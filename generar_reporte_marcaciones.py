import csv
from collections import defaultdict
from datetime import datetime
from pathlib import Path


def parse_fecha(value):
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return None


def parse_hora(value):
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(value, fmt).time()
        except ValueError:
            pass
    return None


def main():
    csv_input = Path("marcacionesSantani3.csv")
    csv_output = Path("reporte_por_ci_y_fecha.csv")

    if not csv_input.exists():
        raise FileNotFoundError(f"No encontré el archivo: {csv_input}")

    with csv_input.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("El CSV está vacío o no tiene cabecera.")

        required = {"CI", "FECHA", "HORA"}
        missing = required - set(reader.fieldnames)
        if missing:
            raise ValueError(f"Faltan columnas requeridas: {sorted(missing)}")

        grouped = defaultdict(list)
        for row in reader:
            ci = str(row.get("CI", "")).strip()
            fecha = parse_fecha(row.get("FECHA"))
            hora = parse_hora(row.get("HORA"))

            if not ci or fecha is None or hora is None:
                continue

            dt = datetime.combine(fecha, hora)
            grouped[(ci, fecha)].append(dt)

    rows = []
    for (ci, fecha), tiempos in sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1])):
        tiempos = sorted(tiempos)
        rows.append({
            "CI": ci,
            "FECHA": fecha.strftime("%d/%m/%Y"),
            "PRIMERA_MARCACION": tiempos[0].strftime("%H:%M:%S"),
            "ULTIMA_MARCACION": tiempos[-1].strftime("%H:%M:%S"),
            "TOTAL_MARCACIONES": len(tiempos),
        })

    with csv_output.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["CI", "FECHA", "PRIMERA_MARCACION", "ULTIMA_MARCACION", "TOTAL_MARCACIONES"],
        )
        writer.writeheader()
        writer.writerows(rows)

    print(f"Reporte generado en: {csv_output}")
    print(f"Registros procesados: {sum(r['TOTAL_MARCACIONES'] for r in rows)}")


if __name__ == "__main__":
    main()

from flask import Flask, render_template, request, send_file, abort
import pandas as pd
from datetime import timedelta
import os
import zipfile
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "files-csv"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FILES_DIR = os.path.join(BASE_DIR, "files")
os.makedirs(FILES_DIR, exist_ok=True)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def add_item(dic, chave, valor):
    if chave not in dic:
        dic[chave] = valor
    else:
        dic[chave] += valor

@app.route("/", methods=["GET", "POST"])
def index():
    zip_name = None
    finished = False

    if request.method == "POST":
        file = request.files["file"]
        if not file:
            return render_template("index.html")

        file_id = str(uuid.uuid4())
        path = os.path.join(UPLOAD_FOLDER, file_id + ".txt")
        file.save(path)
        finished = True

        # ==== PROCESSAMENTO ====
        df = pd.read_csv(path, sep="\t", parse_dates=["DateTime"])
        var_2h = timedelta(hours=2)

        dias_pessoas = {}
        pessoa_horas = {}
        pessoa_dias  = {}
        pessoa_ids   = {}

        nomes = sorted(df["Name"].unique())
        dias  = df["DateTime"].dt.date.unique()

        for dia in dias:
            for nome in nomes:
                if nome not in pessoa_ids:
                    pessoa_ids[nome] = df.loc[df["Name"] == nome, "EnNo"].iloc[0]

                registros = df[
                    (df["DateTime"].dt.date == dia) &
                    (df["Name"] == nome)
                ]

                if len(registros):
                    add_item(dias_pessoas, dia, [nome])
                    add_item(pessoa_dias, nome, 1)

                    if len(registros) == 1:
                        add_item(pessoa_horas, nome, var_2h)
                    else:
                        delta = registros["DateTime"].max() - registros["DateTime"].min()
                        add_item(pessoa_horas, nome, delta)

        run_dir = os.path.join(OUTPUT_FOLDER, file_id)
        os.makedirs(run_dir)

        # ==== DIAS x PESSOAS ====
        max_itens = max(len(p) for p in dias_pessoas.values())
        colunas = ["Data"] + [f"Pessoa {i+1}" for i in range(max_itens)]
        df_out = pd.DataFrame(columns=colunas)

        for dia, pessoas in dias_pessoas.items():
            d = str(dia).split("-")
            data_fmt = f"{d[2]}/{d[1]}/{d[0][2:]}"
            linha = [data_fmt] + pessoas + [None]*(max_itens-len(pessoas))
            df_out.loc[len(df_out)] = linha

        df_out.to_excel(f"{run_dir}/dias-pessoas.xlsx", index=False)

        # ==== PESSOAS x HORAS ====
        df_out2 = pd.DataFrame(columns=["Nome", "ID", "Horas", "Dias presentes"])
        for nome in nomes:
            horas = str(pessoa_horas[nome]).replace("0 days ", "").replace("1 days", "1 dia").replace("days","dias")
            df_out2.loc[len(df_out2)] = [
                nome,
                pessoa_ids[nome],
                horas,
                pessoa_dias[nome]
            ]

        df_out2.to_excel(f"{run_dir}/pessoas-dias-horas.xlsx", index=False)

        zip_name = f"{uuid.uuid4()}.zip"
        zip_path = os.path.join(FILES_DIR, zip_name)

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for file in os.listdir(run_dir):
                z.write(os.path.join(run_dir, file), file)

    return render_template("index.html", zip_name=zip_name, finished=finished)

@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(FILES_DIR, filename)

    if not os.path.exists(path):
        abort(404)

    return send_file(path, as_attachment=True)

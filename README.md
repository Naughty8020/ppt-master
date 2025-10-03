# ppt-master

手順１ 
python -m venv venv

<!-- Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -->


手順２
【1】Windows の場合
コマンドプロンプト（cmd）や PowerShell で：
venv\Scripts\activate

【2】Mac / Linux の場合
ターミナルで：
source venv/bin/activate

成功すると、プロンプトの前に (venv) と表示されます。

手順３
pip install -r requirements.txt
実行python main.py
# 仮想環境に入っている状態でpip freeze > requirements.txt





pip install optimum[openvino]
pip install --upgrade optimum[openvino] optimum-intel

<!-- 
optimum-cli export openvino --model Helsinki-NLP/opus-mt-ja-en --task translation --output openvino_model -->
<!-- 
optimum-cli export openvino --model Helsinki-NLP/opus-mt-ja-en --task translation --weight-format int8 openvino_model -->

optimum-cli export openvino --model facebook/mbart-large-50-many-to-many-mmt --task translation --weight-format int8 openvino_model

Model: データとロジック（例：DB操作、計算など）
View: UI部分（例：GUI、HTML、出力結果）
Controller: ユーザー操作の処理（例：クリック、コマンド実行など）


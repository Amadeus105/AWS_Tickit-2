from flask import Flask, render_template, send_file
import os
import glob
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    """Главная страница с графиками"""
    # Получаем список созданных графиков
    chart_files = glob.glob('charts/*.png')
    charts = []
    
    for chart_file in chart_files:
        chart_name = os.path.basename(chart_file)
        chart_type = chart_name.replace('.png', '').replace('_', ' ').title()
        charts.append({
            'filename': chart_name,
            'name': chart_type,
            'created_time': datetime.fromtimestamp(os.path.getctime(chart_file)).strftime('%Y-%m-%d %H:%M:%S')
        })
    
    # Получаем список Excel файлов
    excel_files = glob.glob('exports/*.xlsx')
    exports = []
    
    for excel_file in excel_files:
        exports.append({
            'filename': os.path.basename(excel_file),
            'size': f"{os.path.getsize(excel_file) / 1024:.1f} KB"
        })
    
    return render_template('index.html', charts=charts, exports=exports)

@app.route('/charts/<filename>')
def serve_chart(filename):
    """Отдача файлов графиков"""
    return send_file(f'charts/{filename}')

@app.route('/exports/<filename>')
def serve_export(filename):
    """Отдача Excel файлов"""
    return send_file(f'exports/{filename}')

@app.route('/run-analysis')
def run_analysis():
    """Запуск анализа и обновление страницы"""
    try:
        from main import AWSTickitAnalyzer
        analyzer = AWSTickitAnalyzer()
        analyzer.run_complete_analysis()
        analyzer.close()
        return "✅ Анализ успешно выполнен! <a href='/'>Вернуться на главную</a>"
    except Exception as e:
        return f"❌ Ошибка: {e} <a href='/'>Вернуться на главную</a>"

if __name__ == '__main__':
    # Создаем необходимые папки
    os.makedirs('charts', exist_ok=True)
    os.makedirs('exports', exist_ok=True)
    os.makedirs('templates', exist_ok=True)
    
    # Запускаем сервер на порту 56777
    print("🚀 Запуск веб-сервера на http://127.0.0.1:56777")
    app.run(host='127.0.0.1', port=56777, debug=True)
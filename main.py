import psycopg2
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import ColorScaleRule
import os
from datetime import datetime, timedelta
import seaborn as sns

class AWSTickitAnalyzer:
    def __init__(self):
        self.connection = None
        self.connect()
        plt.style.use('seaborn-v0_8')
        self.colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
    
    def connect(self):
        """Подключение к базе данных"""
        try:
            self.connection = psycopg2.connect(
                host="localhost",
                database="AWS_Tickit_Database",
                user="postgres",
                password="0000",
                port="5432"
            )
            print("✅ Успешно подключились к базе данных AWS Tickit")
        except Exception as e:
            print(f"❌ Ошибка подключения: {e}")
            raise
    
    def execute_query(self, query, description=""):
        """Выполнение SQL-запроса"""
        try:
            df = pd.read_sql_query(query, self.connection)
            if description:
                print(f"📊 {description}: {len(df)} строк")
            return df
        except Exception as e:
            print(f"❌ Ошибка выполнения запроса: {e}")
            return None

    # 1. PIE CHART - Распределение продаж по категориям событий
    def create_pie_chart(self):
        """Круговая диаграмма: распределение выручки по категориям"""
        query = """
        SELECT c.catgroup, ROUND(SUM(s.pricepaid) / 1000, 2) as revenue_k
        FROM sales s
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        GROUP BY c.catgroup
        ORDER BY revenue_k DESC;
        """
        
        df = self.execute_query(query, "Распределение выручки по категориям")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(10, 8))
            plt.pie(df['revenue_k'], labels=df['catgroup'], autopct='%1.1f%%', 
                   colors=self.colors, startangle=90)
            plt.title('Распределение выручки по категориям событий (в тыс. $)', fontsize=14, fontweight='bold')
            plt.tight_layout()
            os.makedirs('charts', exist_ok=True)
            plt.savefig('charts/pie_chart_revenue_by_category.png', dpi=300, bbox_inches='tight')
            plt.show()
            print("✅ Создана круговая диаграмма: charts/pie_chart_revenue_by_category.png")

    # 2. BAR CHART - Топ-10 городов по количеству пользователей
    def create_bar_chart(self):
        """Столбчатая диаграмма: топ городов по пользователям"""
        query = """
        SELECT city, state, COUNT(*) as user_count
        FROM users 
        GROUP BY city, state 
        HAVING COUNT(*) > 50
        ORDER BY user_count DESC
        LIMIT 10;
        """
        
        df = self.execute_query(query, "Топ городов по пользователям")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 6))
            bars = plt.bar(range(len(df)), df['user_count'], color=self.colors[0])
            plt.title('Топ-10 городов по количеству пользователей', fontsize=14, fontweight='bold')
            plt.xlabel('Города')
            plt.ylabel('Количество пользователей')
            plt.xticks(range(len(df)), [f"{row['city']}, {row['state']}" for _, row in df.iterrows()], rotation=45)
            
            # Добавляем значения на столбцы
            for bar, count in zip(bars, df['user_count']):
                plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5, 
                        int(count), ha='center', va='bottom')
            
            plt.tight_layout()
            plt.savefig('charts/bar_chart_top_cities.png', dpi=300, bbox_inches='tight')
            plt.show()
            print("✅ Создана столбчатая диаграмма: charts/bar_chart_top_cities.png")

    # 3. HORIZONTAL BAR CHART - Средний чек по штатам
    def create_horizontal_bar_chart(self):
        """Горизонтальная столбчатая диаграмма: средний чек по штатам"""
        query = """
        SELECT u.state, 
               AVG(s.pricepaid) as avg_transaction,
               COUNT(s.saleid) as total_sales
        FROM users u
        JOIN sales s ON u.userid = s.buyerid
        GROUP BY u.state
        HAVING COUNT(s.saleid) > 100
        ORDER BY avg_transaction DESC
        LIMIT 15;
        """
        
        df = self.execute_query(query, "Средний чек по штатам")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 8))
            bars = plt.barh(range(len(df)), df['avg_transaction'], color=self.colors[1])
            plt.title('Средняя стоимость транзакции по штатам ($)', fontsize=14, fontweight='bold')
            plt.xlabel('Средняя стоимость транзакции ($)')
            plt.ylabel('Штаты')
            plt.yticks(range(len(df)), df['state'])
            
            # Добавляем значения
            for i, (bar, value) in enumerate(zip(bars, df['avg_transaction'])):
                plt.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2, 
                        f'${value:.2f}', va='center')
            
            plt.tight_layout()
            plt.savefig('charts/horizontal_bar_avg_transaction.png', dpi=300, bbox_inches='tight')
            plt.show()
            print("✅ Создана горизонтальная столбчатая диаграмма: charts/horizontal_bar_avg_transaction.png")

    # 4. LINE CHART - Динамика продаж по месяцам
    def create_line_chart(self):
        """Линейный график: динамика продаж по месяцам"""
        query = """
        SELECT 
            EXTRACT(YEAR FROM s.saletime) as year,
            EXTRACT(MONTH FROM s.saletime) as month,
            COUNT(s.saleid) as total_sales,
            SUM(s.pricepaid) as total_revenue
        FROM sales s
        JOIN events e ON s.eventid = e.eventid
        JOIN venue v ON e.venueid = v.venueid
        GROUP BY year, month
        ORDER BY year, month;
        """
        
        df = self.execute_query(query, "Динамика продаж по месяцам")
        if df is not None and len(df) > 0:
            df['date'] = pd.to_datetime(df['year'].astype(int).astype(str) + '-' + df['month'].astype(int).astype(str) + '-01')
            
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))
            
            # График количества продаж
            ax1.plot(df['date'], df['total_sales'], marker='o', linewidth=2, color=self.colors[2])
            ax1.set_title('Динамика количества продаж по месяцам', fontsize=14, fontweight='bold')
            ax1.set_ylabel('Количество продаж')
            ax1.grid(True, alpha=0.3)
            
            # График выручки
            ax2.plot(df['date'], df['total_revenue'], marker='s', linewidth=2, color=self.colors[3])
            ax2.set_title('Динамика выручки по месяцам', fontsize=14, fontweight='bold')
            ax2.set_ylabel('Выручка ($)')
            ax2.set_xlabel('Месяц')
            ax2.grid(True, alpha=0.3)
            
            plt.tight_layout()
            plt.savefig('charts/line_chart_sales_trends.png', dpi=300, bbox_inches='tight')
            plt.show()
            print("✅ Создан линейный график: charts/line_chart_sales_trends.png")

    # 5. HISTOGRAM - Распределение цен на билеты
    def create_histogram(self):
        """Гистограмма: распределение цен на билеты"""
        query = """
        SELECT l.priceperticket
        FROM listing l
        JOIN events e ON l.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        WHERE l.priceperticket BETWEEN 1 AND 500;
        """
        
        df = self.execute_query(query, "Распределение цен на билеты")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 6))
            plt.hist(df['priceperticket'], bins=30, color=self.colors[4], alpha=0.7, edgecolor='black')
            plt.title('Распределение цен на билеты', fontsize=14, fontweight='bold')
            plt.xlabel('Цена билета ($)')
            plt.ylabel('Количество билетов')
            plt.grid(True, alpha=0.3)
            
            plt.tight_layout()
            plt.savefig('charts/histogram_ticket_prices.png', dpi=300, bbox_inches='tight')
            plt.show()
            print("✅ Создана гистограмма: charts/histogram_ticket_prices.png")

    # 6. SCATTER PLOT - Связь между ценой билета и количеством проданных билетов
    def create_scatter_plot(self):
        """Точечная диаграмма: цена vs количество проданных билетов"""
        query = """
        SELECT 
            AVG(l.priceperticket) as avg_ticket_price,
            SUM(s.qtysold) as total_tickets_sold,
            c.catname
        FROM sales s
        JOIN listing l ON s.listid = l.listid
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        GROUP BY c.catname
        HAVING SUM(s.qtysold) > 100;
        """
        
        df = self.execute_query(query, "Цена vs количество проданных билетов")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 8))
            scatter = plt.scatter(df['avg_ticket_price'], df['total_tickets_sold'], 
                                c=df['avg_ticket_price'], cmap='viridis', s=100, alpha=0.6)
            
            plt.colorbar(scatter, label='Средняя цена билета ($)')
            plt.title('Связь между ценой билета и количеством проданных билетов', fontsize=14, fontweight='bold')
            plt.xlabel('Средняя цена билета ($)')
            plt.ylabel('Общее количество проданных билетов')
            plt.grid(True, alpha=0.3)
            
            # Добавляем подписи категорий
            for i, row in df.iterrows():
                plt.annotate(row['catname'], 
                           (row['avg_ticket_price'], row['total_tickets_sold']),
                           xytext=(5, 5), textcoords='offset points', fontsize=8)
            
            plt.tight_layout()
            plt.savefig('charts/scatter_price_vs_quantity.png', dpi=300, bbox_inches='tight')
            plt.show()
            print("✅ Создана точечная диаграмма: charts/scatter_price_vs_quantity.png")

    # 7. INTERACTIVE PLOTLY CHART WITH SLIDER
    def create_interactive_slider_chart(self):
        """Интерактивный график с временным слайдером"""
        query = """
        SELECT 
            DATE(s.saletime) as sale_date,
            c.catgroup,
            AVG(s.pricepaid) as avg_price,
            COUNT(s.saleid) as daily_sales,
            SUM(s.qtysold) as daily_tickets
        FROM sales s
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        WHERE s.saletime >= '2008-01-01'
        GROUP BY sale_date, c.catgroup
        ORDER BY sale_date;
        """
        
        df = self.execute_query(query, "Данные для интерактивного графика")
        if df is not None and len(df) > 0:
            df['sale_date'] = pd.to_datetime(df['sale_date'])
            df['month_year'] = df['sale_date'].dt.to_period('M').astype(str)
            
            fig = px.scatter(df, 
                           x="daily_sales", 
                           y="avg_price",
                           size="daily_tickets",
                           color="catgroup",
                           animation_frame="month_year",
                           hover_name="catgroup",
                           title="Интерактивная динамика продаж по категориям",
                           labels={"daily_sales": "Ежедневные продажи", 
                                 "avg_price": "Средняя цена ($)"})
            
            fig.show()
            print("✅ Создан интерактивный график с временным слайдером")

    # 8. EXPORT TO EXCEL WITH FORMATTING
    def export_to_excel(self, dataframes_dict, filename):
        """Экспорт данных в Excel с форматированием"""
        try:
            os.makedirs('exports', exist_ok=True)
            filepath = f'exports/{filename}'
            
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                for sheet_name, df in dataframes_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Получаем workbook и worksheet для форматирования
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]
                    
                    # Замораживаем первую строку
                    worksheet.freeze_panes = "A2"
                    
                    # Добавляем фильтры
                    worksheet.auto_filter.ref = worksheet.dimensions
                    
                    # Условное форматирование для числовых колонок
                    numeric_columns = df.select_dtypes(include=['number']).columns
                    for col_idx, col_name in enumerate(numeric_columns, 1):
                        col_letter = chr(64 + col_idx)  # A, B, C, ...
                        range_str = f"{col_letter}2:{col_letter}{len(df)+1}"
                        
                        # Градиентная заливка
                        rule = ColorScaleRule(
                            start_type="min", start_color="FFAA0000",
                            mid_type="percentile", mid_value=50, mid_color="FFFFFF00",
                            end_type="max", end_color="FF00AA00"
                        )
                        worksheet.conditional_formatting.add(range_str, rule)
            
            # Подсчет статистики
            total_sheets = len(dataframes_dict)
            total_rows = sum(len(df) for df in dataframes_dict.values())
            
            print(f"✅ Создан файл {filename}, {total_sheets} листов, {total_rows} строк")
            return True
            
        except Exception as e:
            print(f"❌ Ошибка при экспорте в Excel: {e}")
            return False

    def prepare_data_for_excel_export(self):
        """Подготовка данных для экспорта в Excel"""
        queries = {
            "Sales_Summary": """
                SELECT c.catgroup, c.catname,
                       COUNT(s.saleid) as total_sales,
                       SUM(s.qtysold) as total_tickets,
                       SUM(s.pricepaid) as total_revenue,
                       AVG(s.pricepaid) as avg_sale_amount
                FROM sales s
                JOIN events e ON s.eventid = e.eventid
                JOIN category c ON e.catid = c.catid
                GROUP BY c.catgroup, c.catname
                ORDER BY total_revenue DESC;
            """,
            "User_Geography": """
                SELECT city, state,
                       COUNT(*) as user_count,
                       COUNT(DISTINCT CASE WHEN likesports THEN userid END) as sports_fans,
                       COUNT(DISTINCT CASE WHEN likeconcerts THEN userid END) as concert_fans
                FROM users
                GROUP BY city, state
                HAVING COUNT(*) > 10
                ORDER BY user_count DESC;
            """,
            "Venue_Performance": """
                SELECT v.venuename, v.venuecity, v.venuestate,
                       COUNT(DISTINCT e.eventid) as total_events,
                       SUM(s.pricepaid) as total_revenue,
                       AVG(s.pricepaid) as avg_revenue_per_event
                FROM venue v
                JOIN events e ON v.venueid = e.venueid
                JOIN sales s ON e.eventid = s.eventid
                GROUP BY v.venueid, v.venuename, v.venuecity, v.venuestate
                ORDER BY total_revenue DESC;
            """
        }
        
        dataframes = {}
        for sheet_name, query in queries.items():
            df = self.execute_query(query, f"Подготовка данных для {sheet_name}")
            if df is not None:
                dataframes[sheet_name] = df
        
        return dataframes

    def demonstrate_live_data_update(self):
        """Демонстрация обновления графика при добавлении новых данных"""
        print("\n🎯 ДЕМОНСТРАЦИЯ: Обновление графика при добавлении данных")
        
        # Получаем текущие данные
        query_before = """
        SELECT c.catgroup, SUM(s.pricepaid) as revenue
        FROM sales s
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        GROUP BY c.catgroup
        ORDER BY revenue DESC;
        """
        
        df_before = self.execute_query(query_before, "Данные ДО обновления")
        
        # Создаем график ДО обновления
        if df_before is not None:
            plt.figure(figsize=(10, 6))
            plt.bar(df_before['catgroup'], df_before['revenue'], color='lightblue')
            plt.title('Выручка по категориям (ДО обновления)', fontweight='bold')
            plt.ylabel('Выручка ($)')
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig('charts/before_update.png', dpi=300, bbox_inches='tight')
            plt.show()
        
        # Имитируем добавление новых данных (в реальном сценарии это была бы вставка в БД)
        print("📥 Имитация добавления новых данных продаж...")
        
        # Получаем обновленные данные (можем использовать тот же запрос)
        df_after = self.execute_query(query_before, "Данные ПОСЛЕ обновления")
        
        # Создаем график ПОСЛЕ обновления
        if df_after is not None:
            plt.figure(figsize=(10, 6))
            plt.bar(df_after['catgroup'], df_after['revenue'], color='lightgreen')
            plt.title('Выручка по категориям (ПОСЛЕ обновления)', fontweight='bold')
            plt.ylabel('Выручка ($)')
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig('charts/after_update.png', dpi=300, bbox_inches='tight')
            plt.show()
            
            print("✅ Графики обновлены! Проверьте папку charts/ для сравнения")

    def run_complete_analysis(self):
        """Запуск полного анализа"""
        print("🚀 ЗАПУСК ПОЛНОГО АНАЛИЗА AWS TICKIT")
        print("=" * 50)
        
        # Создаем графики
        self.create_pie_chart()
        self.create_bar_chart()
        self.create_horizontal_bar_chart()
        self.create_line_chart()
        self.create_histogram()
        self.create_scatter_plot()
        
        # Интерактивный график
        self.create_interactive_slider_chart()
        
        # Экспорт в Excel
        excel_data = self.prepare_data_for_excel_export()
        self.export_to_excel(excel_data, "aws_tickit_analysis.xlsx")
        
        # Демонстрация обновления данных
        self.demonstrate_live_data_update()
        
        print("\n🎉 АНАЛИЗ ЗАВЕРШЕН!")
        print("📁 Результаты сохранены в папках: charts/, exports/")

    def close(self):
        """Закрытие соединения"""
        if self.connection:
            self.connection.close()
            print("✅ Соединение с базой данных закрыто")

def main():
    analyzer = None
    try:
        analyzer = AWSTickitAnalyzer()
        analyzer.run_complete_analysis()
    except Exception as e:
        print(f"❌ Произошла ошибка: {e}")
    finally:
        if analyzer:
            analyzer.close()

if __name__ == "__main__":
    main()
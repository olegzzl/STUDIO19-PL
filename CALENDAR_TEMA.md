# 🪂 ЗАПАСНОЙ ПАРАШЮТ: Идеальный календарь бронирования

Этот файл содержит 100% рабочую, отлаженную логику бронирования, которая умеет обходить все подводные камни (CORS, кривые форматы дат от Google и т.д.). Используйте этот код для новых проектов.

---

## 📄 ЧАСТЬ 1: Код для Google Таблиц (Google Apps Script)

Этот код нужно вставить в редакторе скриптов вашей Google Таблицы (Расширения -> Apps Script). 
**Обязательно:** После вставки всегда делайте "Новое развертывание" (New Deployment) от имени вас и с доступом "Для всех" (Anyone), и копируйте НОВЫЙ URL.

```javascript
// НАСТРОЙКИ: Имя листа в таблице, куда будут падать брони (по умолчанию берет первый лист, если такого нет)
const SHEET_NAME = "Bookings";

// Функция для ПРИЕМА данных (когда клиент жмет "Забронировать")
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || 
                  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
                  
    // ЧИТАЕМ JSON: Мы обманываем CORS, присылая text/plain, но внутри это JSON
    const data = JSON.parse(e.postData.contents);
    
    // ЗАПИСЫВАЕМ СТРОКУ (Колонки: A, B, C, D, E, F, G)
    sheet.appendRow([
      new Date(),       // A: Timestamp (когда создана заявка)
      data.name,        // B: Имя клиента
      data.phone,       // C: Телефон
      data.hall,        // D: Выбранный зал
      data.date,        // E: Дата брони (приходит как строка, Google сделает из нее дату)
      data.time,        // F: Время брони
      "Підтверджено"    // G: Статус по умолчанию
    ]);

    // Возвращаем успешный ответ сайту
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Успішно" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Функция для ОТДАЧИ данных (когда сайт загружается и проверяет занятые слоты)
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || 
                SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
                
  const rows = sheet.getDataRange().getValues();
  
  if (rows.length < 2) {
    // Если таблица пустая (только заголовки)
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = [];
  
  // Перебираем все строки, начиная со второй (индекс 1), так как первая - это заголовки
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    // Формируем массив занятых слотов (важно, чтобы порядок колонок совпадал с appendRow)
    data.push({
      timestamp: row[0],
      name: row[1],
      phone: row[2],
      hall: row[3],
      date: row[4], // Google отдаст дату в ISO формате (с буквой T на конце)
      time: row[5], // Google отдаст время в ISO формате (с датой из 1899 года и буквами T/Z)
      status: row[6]
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
```

---

## 💻 ЧАСТЬ 2: Frontend-логика (Javascript для сайта)

Встраивается на ваш сайт. Этот код заменяет обычную форму бронирования. 
**Ключевые фишки:** парсинг кривых форматов дат от старых версий движков Google и отправка формы через `text/plain` во избежание CORS.

```html
<script>
    // ==========================================
    // 1. НАСТРОЙКИ
    // Сюда вставляете URL, который выдал Google Apps Script после деплоя
    // ==========================================
    const GOOGLE_SCRIPT_URL = "ВАШ_URL_ИЗ_APPS_SCRIPT_СЮДА";

    const MONTH_NAMES = [
        "Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень",
        "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень"
    ];

    let bookings = [];
    let selectedHall = "Зал 1";
    
    // Инициализация актуальной даты
    const _today = new Date();
    const _todayStr = `${_today.getFullYear()}-${String(_today.getMonth() + 1).padStart(2, '0')}-${String(_today.getDate()).padStart(2, '0')}`;
    
    let selectedDate = _todayStr; // Стартовая дата (текущий день)
    let selectedTime = null;         
    let currentDate = new Date(_today.getFullYear(), _today.getMonth(), 1); // Текущий месяц календаря

    // Привязка элементов (замените ID на свои, если меняете HTML)
    const calendarContainer = document.getElementById('calendar-container');
    const monthTitle = document.getElementById('month-title');
    const timeSlots = document.querySelectorAll('#time-slots button'); // Кнопки с часами (напр. 10:00, 12:00)

    // ==========================================
    // 2. РЕНДЕР КАЛЕНДАРЯ (Генерация дней месяца)
    // ==========================================
    function renderCalendar() {
        if (!calendarContainer || !monthTitle) return;
        
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth(); 
        
        monthTitle.textContent = `${MONTH_NAMES[month]} ${year}`;
        calendarContainer.innerHTML = '';
        
        // Смещение дней (чтобы 1-е число попадало на правильный день недели)
        const firstDayIndex = new Date(year, month, 1).getDay();
        const firstDayOfWeek = firstDayIndex === 0 ? 7 : firstDayIndex;
        
        const totalDays = new Date(year, month + 1, 0).getDate();
        const prevTotalDays = new Date(year, month, 0).getDate();
        
        // Пустые слоты (дни прошлого месяца)
        for (let j = firstDayOfWeek - 2; j >= 0; j--) {
            const emptyCell = document.createElement('div');
            emptyCell.className = "text-gray-300 flex items-center justify-center";
            emptyCell.textContent = prevTotalDays - j;
            calendarContainer.appendChild(emptyCell);
        }
        
        // Основные дни месяца
        for (let i = 1; i <= totalDays; i++) {
            const btn = document.createElement('button');
            const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(i).padStart(2, '0')}`;
            
            btn.textContent = i;
            // Логика выделения активного дня
            if (dateStr === selectedDate) {
                btn.className = "bg-black text-white day-btn"; // Стили активного дня
            } else {
                btn.className = "border border-gray-200 day-btn"; // Стили обычного дня
            }
            
            btn.addEventListener('click', () => {
                selectedDate = dateStr;
                // Снимаем выделение со всех и ставим на текущую
                document.querySelectorAll('.day-btn').forEach(b => b.className = "border border-gray-200 day-btn");
                btn.className = "bg-black text-white day-btn";
                
                updateAvailability(); // Проверяем свободные слоты для нового дня
            });
            calendarContainer.appendChild(btn);
        }
    }

    // ==========================================
    // 3. ПОЛУЧЕНИЕ БРОНИРОВАНИЙ (С фиксом багов Google)
    // ==========================================
    async function loadBookings() {
        if (!GOOGLE_SCRIPT_URL) return;
        try {
            // Добавляем ?_t=... чтобы сбросить кэш браузера!
            const response = await fetch(`${GOOGLE_SCRIPT_URL}?_t=${new Date().getTime()}`);
            if (response.ok) {
                const rawData = await response.json();
                
                // МЕГА-ФИКС ДЛЯ GOOGLE SHEETS: 
                // Google Sheets часто отдает даты как ISO строки, а время как 1899-12-30T11:57:56.000Z
                // Этот код нормализует их обратно в YYYY-MM-DD и HH:mm
                bookings = rawData.map(b => {
                    let nDate = b.date;
                    if (b.date && String(b.date).includes("T")) {
                        const d = new Date(b.date);
                        if (!isNaN(d)) nDate = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                    }
                    let nTime = b.time;
                    if (b.time && (String(b.time).includes("GMT") || String(b.time).includes("T") || String(b.time).includes("Z"))) {
                        const t = new Date(b.time);
                        if (!isNaN(t)) {
                            let h = t.getHours(); let m = t.getMinutes();
                            if (m > 45) { h = (h + 1) % 24; m = 0; }
                            else if (m > 15 && m < 45) { m = 30; }
                            else { m = 0; }
                            nTime = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
                        }
                    }
                    return { ...b, date: nDate, time: nTime };
                });
            }
        } catch (err) {
            console.error("Ошибка сети:", err);
        }
        updateAvailability();
    }

    // ==========================================
    // 4. БЛОКИРОВКА ЗАНЯТОГО ВРЕМЕНИ
    // ==========================================
    function updateAvailability() {
        timeSlots.forEach(slot => {
            const slotTime = slot.textContent.trim();
            // Проверяем, есть ли такое бронирование в базе
            const isBooked = bookings.some(b => 
                b.date === selectedDate && 
                b.time === slotTime && 
                b.hall === selectedHall &&
                String(b.status || "Підтверджено").toLowerCase() !== "скасовано"
            );
            
            if (isBooked) {
                // Если занято -> делаем зачеркнутым, снимаем клики
                slot.className = "bg-gray-100 text-gray-300 pointer-events-none line-through";
                if (selectedTime === slotTime) selectedTime = null;
            } else {
                // Если свободно -> делаем обычным
                if (selectedTime === slotTime) {
                    slot.className = "bg-black text-white slot-btn";
                } else {
                    slot.className = "border border-black slot-btn";
                }
            }
        });
    }

    // ==========================================
    // 5. ОТПРАВКА ДАННЫХ В GOOGLE (Обход CORS)
    // ==========================================
    document.getElementById('booking-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        
        if (!selectedTime) {
            alert("Будь ласка, оберіть доступний час!"); return;
        }
        
        const payload = {
            name: document.getElementById('client-name').value,
            phone: document.getElementById('client-phone').value,
            hall: selectedHall,
            date: selectedDate,
            time: selectedTime
        };
        
        try {
            // ФИКС CORS: Отправляем JSON, но под видом text/plain
            const response = await fetch(GOOGLE_SCRIPT_URL, {
                method: 'POST',
                body: JSON.stringify(payload),
                headers: { 'Content-Type': 'text/plain;charset=utf-8' }
            });
            
            const result = await response.json();
            if (result.status === "success") {
                // Добавляем бронь в память чтобы сразу заблокировать слот без перезагрузки
                bookings.push({ date: selectedDate, time: selectedTime, hall: selectedHall, status: "Підтверджено" });
                updateAvailability();
                alert("Успішно заброньовано!");
            }
        } catch (err) {
            alert("Помилка відправки");
        }
    });

    // Запускаем всё
    renderCalendar();
    loadBookings();
</script>
```

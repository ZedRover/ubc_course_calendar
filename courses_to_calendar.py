import pandas as pd
from icalendar import Calendar, Event, vText
from datetime import datetime, timedelta
import pytz
import re
import warnings

# Ignore openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Set Vancouver time zone
vancouver_tz = pytz.timezone('America/Vancouver')

# Load Excel file
file_path = 'View_My_Courses.xlsx'
df = pd.read_excel(file_path, header=2)  # Set header=2 to skip the first two rows

# Create calendar
cal = Calendar()
cal.add('prodid', '-//UBC Course Calendar//EN')
cal.add('version', '2.0')
cal.add('x-wr-timezone', 'America/Vancouver')

# 解析 Meeting Pattern
def parse_meeting_pattern(pattern):
    if pd.isna(pattern):  # 检查是否为空
        return []
    patterns = pattern.split('\n')
    result = []
    for p in patterns:
        match = re.match(r'(\d{4}-\d{2}-\d{2}) - (\d{4}-\d{2}-\d{2})\s*\|\s*(\w+(?:\s+\w+)*)\s*\|\s*(\d{1,2}:\d{2}\s*[ap]\.?m\.?)\s*-\s*(\d{1,2}:\d{2}\s*[ap]\.?m\.?)\s*\|\s*(.+)', p.strip())
        if match:
            start_date, end_date, days, start_time, end_time, location = match.groups()
            result.append((start_date, end_date, days, start_time.strip(), end_time.strip(), location))
    return result

# 转换时间格式
def convert_time(time_str):
    # 移除所有的点号和多余的空格
    time_str = time_str.replace('.', '').strip()
    
    # 使用正则表达式来匹配时间格式
    match = re.match(r'(\d{1,2}):(\d{2})\s*(am|pm)', time_str, re.IGNORECASE)
    if match:
        hour, minute, period = match.groups()
        hour = int(hour)
        if period.lower() == 'pm' and hour != 12:
            hour += 12
        elif period.lower() == 'am' and hour == 12:
            hour = 0
        return timedelta(hours=hour, minutes=int(minute))
    else:
        raise ValueError(f"无法解析时间字符串: {time_str}")

# 获取星期几的缩写
def get_weekdays(days):
    day_map = {'Mon': 'MO', 'Tue': 'TU', 'Wed': 'WE', 'Thu': 'TH', 'Fri': 'FR', 'Sat': 'SA', 'Sun': 'SU'}
    return [day_map[day] for day in days.split()]

# 遍历 DataFrame 的行
for _, row in df.iterrows():
    patterns = parse_meeting_pattern(row['Meeting Patterns'])
    if not patterns:  # 如果 Meeting Pattern 为空，跳过这一行
        continue
    for start_date, end_date, days, start_time, end_time, location in patterns:
        start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date, '%Y-%m-%d').date()
        start_time = convert_time(start_time)
        end_time = convert_time(end_time)
        weekdays = get_weekdays(days)

        event = Event()
        event.add('summary', row['Course Listing'])
        event_start = vancouver_tz.localize(datetime.combine(start_date, datetime.min.time()) + start_time)
        event_end = vancouver_tz.localize(datetime.combine(start_date, datetime.min.time()) + end_time)
        event.add('dtstart', event_start)
        event.add('dtend', event_end)
        event.add('description', f"Instructor: {row['Instructor']}\nSection: {row['Section']}")
        event.add('location', vText(f"{row['Delivery Mode']} | {location}"))

        # Add weekly recurrence rule
        rrule = {
            'freq': 'weekly',
            'byday': weekdays,
            'until': vancouver_tz.localize(datetime.combine(end_date, datetime.max.time()))
        }
        event.add('rrule', rrule)

        cal.add_component(event)

# Write calendar to file
calendar_file = 'calendar.ics'
with open(calendar_file, 'wb') as f:
    f.write(cal.to_ical())

print(f"Calendar file '{calendar_file}' has been created.")
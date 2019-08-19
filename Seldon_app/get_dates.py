from datetime import datetime, date, time

TIME_WINDOW = 7

def get_dates(time_window=TIME_WINDOW):
      '''
      counts time-window from tomorrow - 7 days 
      returns 2 tuples (year, month, day)
      '''
      days_in_month = 30
      today = datetime.today()
      current_year = today.year
      current_month = today.month
      
      current_day = int()
      if today.day < 30:
            current_day = today.day + 1
      else:
            current_day = 1
            current_month += 1

      day_from = int()
      month_from = int()
      year_from = int()
            
      if (current_day - time_window) < 0:
            day_from = days_in_month - (time_window - current_day)
            month_from -= 1
      elif (current_day - time_window) == 0:
            day_from = 1
      else:
            day_from = current_day - time_window

      if (month_from + current_month) == 0:
            month_from = 12
            year_from = current_year - 1
      else:
            month_from = month_from + current_month
            year_from = current_year

      current_year = str(current_year)
      current_month = str(current_month)
      current_day = str(current_day)
      
      if len(str(day_from)) < 2:
            day_from = f'0{day_from}'
      else:
            day_from = str(day_from)
            
      if len(str(month_from)) < 2:
            month_from = f'0{month_from}'
      else:
            month_from = str(month_from)
            
      year_from = str(year_from)

      date_from = (year_from, month_from, day_from)
      date_to = (current_year, current_month, current_day)
      
      return date_from, date_to




























































      
      

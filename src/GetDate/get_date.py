from datetime import datetime , timedelta

class get_period:

    ## ONTEM:##################################################################

    def yesterday_dmy_period(self):
        return (datetime.now() - timedelta(days=1)).strftime('%d.%m.%Y')
    
    def yesterday_dmy_bar(self):
        return (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
    
    def yesterday_dmy_dash(self):
        return (datetime.now() - timedelta(days=1)).strftime('%d-%m-%Y')

    def yesterday_ymd_period(self):
        return (datetime.now() - timedelta(days=1)).strftime('%Y.%m.%d')

    def yesterday_ymd_bar(self):
        return (datetime.now() - timedelta(days=1)).strftime('%Y/%m/%d')
    
    ## HOJE:###################################################################

    def today_dmy_period(self):
        return (datetime.now()).strftime('%d.%m.%Y')
    
    def today_dmy_bar(self):
        return (datetime.now()).strftime('%d/%m/%Y')
    
    def today_dmy_dash(self):
        return (datetime.now()).strftime('%d-%m-%Y')

    def today_ymd_period(self):
        return (datetime.now()).strftime('%Y.%m.%d')

    def today_ymd_bar(self):
        return (datetime.now()).strftime('%Y/%m/%d')
    
    ## QUALQUER DIA ANTERIOR:###################################################

    def previous_date_dmy_period(self, days=1):
        return (datetime.now() - timedelta(days)).strftime('%d.%m.%Y')
    
    def previous_date_dmy_bar(self, days=1):
        return (datetime.now() - timedelta(days)).strftime('%d/%m/%Y')
    
    def previous_date_dmy_dash(self, days=1):
        return (datetime.now() - timedelta(days)).strftime('%d-%m-%Y')

    def previous_date_ymd_period(self, days=1):
        return (datetime.now() - timedelta(days)).strftime('%Y.%m.%d')

    def previous_date_ymd_bar(self, days=1):
        return (datetime.now() - timedelta(days)).strftime('%Y/%m/%d')

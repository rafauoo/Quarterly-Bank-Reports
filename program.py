import xlsxwriter

months = [
    "Styczeń",
    "Luty",
    "Marzec",
    "Kwiecień",
    "Maj",
    "Czerwiec",
    "Lipiec",
    "Sierpień",
    "Wrzesień",
    "Październik",
    "Listopad",
    "Grudzień"
]

class Program:
    def __init__(self, file_name, year) -> None:
        self._workbook = xlsxwriter.Workbook(file_name)
        self._worksheet = self._workbook.add_worksheet()
        self._options = {}
        self._year = year
        self._columns = 0
        self._sums = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        self._options["bold"] = self._workbook.add_format({'bold': True})
        self.values()

    def values(self):
        green = self._workbook.add_format({
            'fg_color': '548235'})
        blue = self._workbook.add_format({
            'fg_color': '2f75b5'})
        pink = self._workbook.add_format({
            'fg_color': 'ca3a90'})
        for i in range(6):
            k = 0 if i <= 2 else 1
            self._worksheet.write(i * 6 + 4+k, self._columns, "100", green)
            self._worksheet.write(i * 6 + 5+k, self._columns, "50", blue)
            self._worksheet.write(i * 6 + 6+k, self._columns, "20", pink)
        self._columns += 1
    
    def end(self):
        self.values()
        self._workbook.close()
    
    def write(self, row, column, value, option=None):
        self._worksheet.write(row, column, value, self._options[option])
    
    def write_month(self, dates, month):
        sum_month_100 = 0
        sum_month_50 = 0
        sum_month_20 = 0
        sum_month_100o = 0
        sum_month_50o = 0
        sum_month_20o = 0
        sum_month_100n = 0
        sum_month_50n = 0
        sum_month_20n = 0
        sum_month_100on = 0
        sum_month_50on = 0
        sum_month_20on = 0
        merge_yellow = self._workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'yellow'})
        merge_orange = self._workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'orange'})
        length = len(dates)
        green = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'c4d79b'})
        blue = self._workbook.add_format({
                        'border': 1,
            'fg_color': '9BB7D9'})
        pink = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'da8ac1'})
        lgreen = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'd8e4bc'})
        lblue = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'b8cce4'})
        lpink = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'eec8e2'})
        org = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'f8ab6c'})
        lorg = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'fce0c8'})
        mo = months[month - 1]
        self._worksheet.merge_range(0, self._columns, 0, self._columns + 2 * length + 2, f"{mo} {self._year}", cell_format=merge_yellow)
        self._worksheet.merge_range(1, self._columns, 1, self._columns + 2 * length + 2, f"Załadunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(7, self._columns, 7, self._columns + 2 * length + 2, f"Rozładunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(13, self._columns, 13, self._columns + 2 * length + 2, f"Załadunek - Rozładunek", cell_format=merge_orange)
        self._worksheet.merge_range(19, self._columns, 19, self._columns + 2 * length + 2, f"{mo} {self._year}", cell_format=merge_yellow)
        self._worksheet.merge_range(20, self._columns, 20, self._columns + 2 * length + 2, f"Załadunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(26, self._columns, 26, self._columns + 2 * length + 2, f"Rozładunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(32, self._columns, 32, self._columns + 2 * length + 2, f"Załadunek - Rozładunek", cell_format=merge_orange)
        for date_data in dates:
            date, data = date_data
            for i in range(6):
                k = 0 if i <= 2 else 1
                self._worksheet.merge_range(i*6+2+k, self._columns, i*6+2+k, self._columns + 1, f"{date}.{month:02}", cell_format=merge_orange)
            self._worksheet.write(3, self._columns, "NBP", org)
            self._worksheet.write(4, self._columns, data[0], green)
            sum_month_100n += data[0]
            self._worksheet.write(5, self._columns, data[1], blue)
            sum_month_50n += data[1]
            self._worksheet.write(6, self._columns, data[2], pink)
            sum_month_20n += data[2]
            self._worksheet.write(3, self._columns + 1, "Własne", lorg)
            self._worksheet.write(4, self._columns + 1, data[3], lgreen)
            sum_month_100 += data[3]
            self._worksheet.write(5, self._columns + 1, data[4], lblue)
            sum_month_50 += data[4]
            self._worksheet.write(6, self._columns + 1, data[5], lpink)
            sum_month_20 += data[5]
            self._worksheet.write(9, self._columns, "NBP", org)
            self._worksheet.write(10, self._columns, data[6], green)
            sum_month_100on += data[6]
            self._worksheet.write(11, self._columns, data[7], blue)
            sum_month_50on += data[7]
            self._worksheet.write(12, self._columns, data[8], pink)
            sum_month_20on += data[8]
            self._worksheet.write(9, self._columns + 1, "Własne", lorg)
            self._worksheet.write(10, self._columns + 1, data[9], lgreen)
            sum_month_100o += data[9]
            self._worksheet.write(11, self._columns + 1, data[10], lblue)
            sum_month_50o += data[10]
            self._worksheet.write(12, self._columns + 1, data[11], lpink)
            sum_month_20o += data[11]
            self._worksheet.write(15, self._columns, "NBP", org)
            self._worksheet.write(16, self._columns, data[0] - data[6], green)
            self._worksheet.write(17, self._columns, data[1] - data[7], blue)
            self._worksheet.write(18, self._columns, data[2] - data[8], pink)
            self._worksheet.write(15, self._columns + 1, "Własne", lorg)
            self._worksheet.write(16, self._columns + 1, data[3] - data[9], lgreen)
            self._worksheet.write(17, self._columns + 1, data[4] - data[10], lblue)
            self._worksheet.write(18, self._columns + 1, data[5] - data[11], lpink)
            # zł
            self._worksheet.write(22, self._columns, "NBP", org)
            self._worksheet.write(23, self._columns, data[0] * 100, green)
            self._worksheet.write(24, self._columns, data[1] * 50, blue)
            self._worksheet.write(25, self._columns, data[2] * 20, pink)
            self._worksheet.write(22, self._columns + 1, "Własne", lorg)
            self._worksheet.write(23, self._columns + 1, data[3] * 100, lgreen)
            self._worksheet.write(24, self._columns + 1, data[4] * 50, lblue)
            self._worksheet.write(25, self._columns + 1, data[5] * 20, lpink)
            self._worksheet.write(28, self._columns, "NBP", org)
            self._worksheet.write(29, self._columns, data[6] * 100, green)
            self._worksheet.write(30, self._columns, data[7] * 50, blue)
            self._worksheet.write(31, self._columns, data[8] * 20, pink)
            self._worksheet.write(28, self._columns + 1, "Własne", lorg)
            self._worksheet.write(29, self._columns + 1, data[9] * 100, lgreen)
            self._worksheet.write(30, self._columns + 1, data[10] * 50, lblue)
            self._worksheet.write(31, self._columns + 1, data[11] * 20, lpink)
            self._worksheet.write(34, self._columns, "NBP", org)
            self._worksheet.write(35, self._columns, (data[0] - data[6]) * 100, green)
            self._worksheet.write(36, self._columns, (data[1] - data[7]) * 50, blue)
            self._worksheet.write(37, self._columns, (data[2] - data[8]) * 20, pink)
            self._worksheet.write(34, self._columns + 1, "Własne", lorg)
            self._worksheet.write(35, self._columns + 1, (data[3] - data[9])*100, lgreen)
            self._worksheet.write(36, self._columns + 1, (data[4] - data[10])*50, lblue)
            self._worksheet.write(37, self._columns + 1, (data[5] - data[11])*20, lpink)
            self._columns += 2
#===================SUM COLUMNS==========================
        self._worksheet.merge_range(2, self._columns, 2, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(8, self._columns, 8, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(14, self._columns, 14, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(21, self._columns, 21, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(27, self._columns, 27, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(33, self._columns, 33, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.write(3, self._columns, "NBP", org)
        self._worksheet.write(4, self._columns, sum_month_100n, green)
        self._sums[0] += sum_month_100n
        self._worksheet.write(5, self._columns, sum_month_50n, blue)
        self._sums[1] += sum_month_50n
        self._worksheet.write(6, self._columns, sum_month_20n, pink)
        self._sums[2] += sum_month_20n
        self._worksheet.write(3, self._columns + 1, "Własne", lorg)
        self._worksheet.write(4, self._columns + 1, sum_month_100, lgreen)
        self._sums[3] += sum_month_100
        self._worksheet.write(5, self._columns + 1, sum_month_50, lblue)
        self._sums[4] += sum_month_50
        self._worksheet.write(6, self._columns + 1, sum_month_20, lpink)
        self._sums[5] += sum_month_20
        self._worksheet.write(9, self._columns, "NBP", org)
        self._worksheet.write(10, self._columns, sum_month_100on, green)
        self._sums[6] += sum_month_100on
        self._worksheet.write(11, self._columns, sum_month_50on, blue)
        self._sums[7] += sum_month_50on
        self._worksheet.write(12, self._columns, sum_month_20on, pink)
        self._sums[8] += sum_month_20on
        self._worksheet.write(9, self._columns + 1, "Własne", lorg)
        self._worksheet.write(10, self._columns + 1, sum_month_100o, lgreen)
        self._sums[9] += sum_month_100o
        self._worksheet.write(11, self._columns + 1, sum_month_50o, lblue)
        self._sums[10] += sum_month_50o
        self._worksheet.write(12, self._columns + 1, sum_month_20o, lpink)
        self._sums[11] += sum_month_20o
        self._worksheet.write(15, self._columns, "NBP", org)
        self._worksheet.write(16, self._columns, sum_month_100n - sum_month_100on, green)
        self._worksheet.write(17, self._columns, sum_month_50n - sum_month_50on, blue)
        self._worksheet.write(18, self._columns, sum_month_20n - sum_month_20on, pink)
        self._worksheet.write(15, self._columns + 1, "Własne", lorg)
        self._worksheet.write(16, self._columns + 1, sum_month_100 - sum_month_100o, lgreen)
        self._worksheet.write(17, self._columns + 1, sum_month_50 - sum_month_50o, lblue)
        self._worksheet.write(18, self._columns + 1, sum_month_20 - sum_month_20o, lpink)
        # zł
        self._worksheet.write(22, self._columns, "NBP", org)
        self._worksheet.write(23, self._columns, sum_month_100n * 100, green)
        self._worksheet.write(24, self._columns, sum_month_50n * 50, blue)
        self._worksheet.write(25, self._columns, sum_month_20n * 20, pink)
        self._worksheet.write(22, self._columns + 1, "Własne", lorg)
        self._worksheet.write(23, self._columns + 1, sum_month_100 * 100, lgreen)
        self._worksheet.write(24, self._columns + 1, sum_month_50 * 50, lblue)
        self._worksheet.write(25, self._columns + 1, sum_month_20 * 20, lpink)
        self._worksheet.write(28, self._columns, "NBP", org)
        self._worksheet.write(29, self._columns, sum_month_100on * 100, green)
        self._worksheet.write(30, self._columns, sum_month_50on * 50, blue)
        self._worksheet.write(31, self._columns, sum_month_20on * 20, pink)
        self._worksheet.write(28, self._columns + 1, "Własne", lorg)
        self._worksheet.write(29, self._columns + 1, sum_month_100o * 100, lgreen)
        self._worksheet.write(30, self._columns + 1, sum_month_50o * 50, lblue)
        self._worksheet.write(31, self._columns + 1, sum_month_20o * 20, lpink)
        self._worksheet.write(34, self._columns, "NBP", org)
        self._worksheet.write(35, self._columns, (sum_month_100n - sum_month_100on) * 100, green)
        self._worksheet.write(36, self._columns, (sum_month_50n - sum_month_50on) * 50, blue)
        self._worksheet.write(37, self._columns, (sum_month_20n - sum_month_20on) * 20, pink)
        self._worksheet.write(34, self._columns + 1, "Własne", lorg)
        self._worksheet.write(35, self._columns + 1, (sum_month_100 - sum_month_100o)*100, lgreen)
        self._worksheet.write(36, self._columns + 1, (sum_month_50 - sum_month_50o)*50, lblue)
        self._worksheet.write(37, self._columns + 1, (sum_month_20 - sum_month_20o)*20, lpink)
        self._columns += 2
#==========SUM ALL===================
        self._worksheet.write(3, self._columns, "RAZEM", org)
        self._worksheet.write(4, self._columns, sum_month_100n + sum_month_100, green)
        self._worksheet.write(5, self._columns, sum_month_50n + sum_month_50, blue)
        self._worksheet.write(6, self._columns, sum_month_20n + sum_month_20, pink)
        self._worksheet.write(9, self._columns, "RAZEM", org)
        self._worksheet.write(10, self._columns, sum_month_100on + sum_month_100o, green)
        self._worksheet.write(11, self._columns, sum_month_50on + sum_month_50o, blue)
        self._worksheet.write(12, self._columns, sum_month_20on + sum_month_20o, pink)
        self._worksheet.write(15, self._columns, "RAZEM", org)
        self._worksheet.write(16, self._columns, sum_month_100n + sum_month_100 - sum_month_100on - sum_month_100o, green)
        self._worksheet.write(17, self._columns, sum_month_50n + sum_month_50 - sum_month_50on - sum_month_50o, blue)
        self._worksheet.write(18, self._columns, sum_month_20n + sum_month_20 - sum_month_20on - sum_month_20o, pink)
        # zł
        self._worksheet.write(22, self._columns, "RAZEM", org)
        self._worksheet.write(23, self._columns, (sum_month_100n + sum_month_100)*100, green)
        self._worksheet.write(24, self._columns, (sum_month_50n + sum_month_50)*50, blue)
        self._worksheet.write(25, self._columns, (sum_month_20n + sum_month_20)*20, pink)
        self._worksheet.write(28, self._columns, "RAZEM", org)
        self._worksheet.write(29, self._columns, (sum_month_100on + sum_month_100o)*100, green)
        self._worksheet.write(30, self._columns, (sum_month_50on + sum_month_50o)*50, blue)
        self._worksheet.write(31, self._columns, (sum_month_20on + sum_month_20o)*20, pink)
        self._worksheet.write(34, self._columns, "RAZEM", org)
        self._worksheet.write(35, self._columns, (sum_month_100n + sum_month_100 - sum_month_100on - sum_month_100o)*100, green)
        self._worksheet.write(36, self._columns, (sum_month_50n + sum_month_50 - sum_month_50on - sum_month_50o)*50, blue)
        self._worksheet.write(37, self._columns, (sum_month_20n + sum_month_20 - sum_month_20on - sum_month_20o)*20, pink)
        black = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'black'})
        self._worksheet.set_column(self._columns + 1, self._columns + 1, 0.2, black)
        self._columns += 2


    def sum_all(self):
        merge_orange = self._workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'orange'})
        merge_yellow = self._workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'yellow'})
        green = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'c4d79b'})
        blue = self._workbook.add_format({
                        'border': 1,
            'fg_color': '9BB7D9'})
        pink = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'da8ac1'})
        lgreen = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'd8e4bc'})
        lblue = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'b8cce4'})
        lpink = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'eec8e2'})
        org = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'f8ab6c'})
        lorg = self._workbook.add_format({
                        'border': 1,
            'fg_color': 'fce0c8'})
        self._worksheet.merge_range(0, self._columns, 0, self._columns + 2, f"RAZEM", cell_format=merge_yellow)
        self._worksheet.merge_range(1, self._columns, 1, self._columns + 2, f"Załadunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(7, self._columns, 7, self._columns + 2, f"Rozładunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(13, self._columns, 13, self._columns + 2, f"Załadunek - Rozładunek", cell_format=merge_orange)
        self._worksheet.merge_range(19, self._columns, 19, self._columns + 2, f"RAZEM", cell_format=merge_yellow)
        self._worksheet.merge_range(20, self._columns, 20, self._columns + 2, f"Załadunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(26, self._columns, 26, self._columns + 2, f"Rozładunek bankomatu", cell_format=merge_orange)
        self._worksheet.merge_range(32, self._columns, 32, self._columns + 2, f"Załadunek - Rozładunek", cell_format=merge_orange)
#===================SUM COLUMNS==========================
        self._worksheet.merge_range(2, self._columns, 2, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(8, self._columns, 8, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(14, self._columns, 14, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(21, self._columns, 21, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(27, self._columns, 27, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.merge_range(33, self._columns, 33, self._columns + 2, f"SUMA", cell_format=merge_orange)
        self._worksheet.write(3, self._columns, "NBP", org)
        self._worksheet.write(4, self._columns, self._sums[0], green)
        self._worksheet.write(5, self._columns, self._sums[1], blue)
        self._worksheet.write(6, self._columns, self._sums[2], pink)
        self._worksheet.write(3, self._columns + 1, "Własne", lorg)
        self._worksheet.write(4, self._columns + 1, self._sums[3], lgreen)
        self._worksheet.write(5, self._columns + 1, self._sums[4], lblue)
        self._worksheet.write(6, self._columns + 1, self._sums[5], lpink)
        self._worksheet.write(9, self._columns, "NBP", org)
        self._worksheet.write(10, self._columns, self._sums[6], green)
        self._worksheet.write(11, self._columns, self._sums[7], blue)
        self._worksheet.write(12, self._columns, self._sums[8], pink)
        self._worksheet.write(9, self._columns + 1, "Własne", lorg)
        self._worksheet.write(10, self._columns + 1, self._sums[9], lgreen)
        self._worksheet.write(11, self._columns + 1, self._sums[10], lblue)
        self._worksheet.write(12, self._columns + 1, self._sums[11], lpink)
        self._worksheet.write(15, self._columns, "NBP", org)
        self._worksheet.write(16, self._columns, self._sums[0] - self._sums[6], green)
        self._worksheet.write(17, self._columns, self._sums[1] - self._sums[7], blue)
        self._worksheet.write(18, self._columns, self._sums[2] - self._sums[8], pink)
        self._worksheet.write(15, self._columns + 1, "Własne", lorg)
        self._worksheet.write(16, self._columns + 1, self._sums[3] - self._sums[9], lgreen)
        self._worksheet.write(17, self._columns + 1, self._sums[4] - self._sums[10], lblue)
        self._worksheet.write(18, self._columns + 1, self._sums[5] - self._sums[11], lpink)
        # zł
        self._worksheet.write(22, self._columns, "NBP", org)
        self._worksheet.write(23, self._columns, self._sums[0] * 100, green)
        self._worksheet.write(24, self._columns, self._sums[1] * 50, blue)
        self._worksheet.write(25, self._columns, self._sums[2] * 20, pink)
        self._worksheet.write(22, self._columns + 1, "Własne", lorg)
        self._worksheet.write(23, self._columns + 1, self._sums[3] * 100, lgreen)
        self._worksheet.write(24, self._columns + 1, self._sums[4] * 50, lblue)
        self._worksheet.write(25, self._columns + 1, self._sums[5] * 20, lpink)
        self._worksheet.write(28, self._columns, "NBP", org)
        self._worksheet.write(29, self._columns, self._sums[6] * 100, green)
        self._worksheet.write(30, self._columns, self._sums[7] * 50, blue)
        self._worksheet.write(31, self._columns, self._sums[8] * 20, pink)
        self._worksheet.write(28, self._columns + 1, "Własne", lorg)
        self._worksheet.write(29, self._columns + 1, self._sums[9] * 100, lgreen)
        self._worksheet.write(30, self._columns + 1, self._sums[10] * 50, lblue)
        self._worksheet.write(31, self._columns + 1, self._sums[11] * 20, lpink)
        self._worksheet.write(34, self._columns, "NBP", org)
        self._worksheet.write(35, self._columns, (self._sums[0] - self._sums[6]) * 100, green)
        self._worksheet.write(36, self._columns, (self._sums[1] - self._sums[7]) * 50, blue)
        self._worksheet.write(37, self._columns, (self._sums[2] - self._sums[8]) * 20, pink)
        self._worksheet.write(34, self._columns + 1, "Własne", lorg)
        self._worksheet.write(35, self._columns + 1, (self._sums[3] - self._sums[9])*100, lgreen)
        self._worksheet.write(36, self._columns + 1, (self._sums[4] - self._sums[10])*50, lblue)
        self._worksheet.write(37, self._columns + 1, (self._sums[5] - self._sums[11])*20, lpink)
        self._columns += 2
#==========SUM ALL===================
        self._worksheet.write(3, self._columns, "RAZEM", org)
        self._worksheet.write(4, self._columns, self._sums[0] + self._sums[3], green)
        self._worksheet.write(5, self._columns, self._sums[1] + self._sums[4], blue)
        self._worksheet.write(6, self._columns, self._sums[2] + self._sums[5], pink)
        self._worksheet.write(9, self._columns, "RAZEM", org)
        self._worksheet.write(10, self._columns, self._sums[6] + self._sums[9], green)
        self._worksheet.write(11, self._columns, self._sums[7] + self._sums[10], blue)
        self._worksheet.write(12, self._columns, self._sums[8] + self._sums[11], pink)
        self._worksheet.write(15, self._columns, "RAZEM", org)
        self._worksheet.write(16, self._columns, self._sums[0] + self._sums[3] - self._sums[6] - self._sums[9], green)
        self._worksheet.write(17, self._columns, self._sums[1] + self._sums[4] - self._sums[7] - self._sums[10], blue)
        self._worksheet.write(18, self._columns, self._sums[2] + self._sums[5] - self._sums[8] - self._sums[11], pink)
        # zł
        self._worksheet.write(22, self._columns, "RAZEM", org)
        self._worksheet.write(23, self._columns, (self._sums[0] + self._sums[3])*100, green)
        self._worksheet.write(24, self._columns, (self._sums[1] + self._sums[4])*50, blue)
        self._worksheet.write(25, self._columns, (self._sums[2] + self._sums[5])*20, pink)
        self._worksheet.write(28, self._columns, "RAZEM", org)
        self._worksheet.write(29, self._columns, (self._sums[6] + self._sums[9])*100, green)
        self._worksheet.write(30, self._columns, (self._sums[7] + self._sums[10])*50, blue)
        self._worksheet.write(31, self._columns, (self._sums[8] + self._sums[11])*20, pink)
        self._worksheet.write(34, self._columns, "RAZEM", org)
        self._worksheet.write(35, self._columns, (self._sums[0] + self._sums[3] - self._sums[6] - self._sums[9])*100, green)
        self._worksheet.write(36, self._columns, (self._sums[1] + self._sums[4] - self._sums[7] - self._sums[10])*50, blue)
        self._worksheet.write(37, self._columns, (self._sums[2] + self._sums[5] - self._sums[8] - self._sums[11])*20, pink)
        self._columns += 1
from program import Program

class ConsoleInterface:
    def __init__(self) -> None:
        file_name, year = self.get_file_name()
        program = Program(file_name, year)
        dates, month = self.get_month_data()
        dates = sorted(dates)
        program.write_month(dates, month)
        dates, month = self.get_month_data()
        dates = sorted(dates)
        program.write_month(dates, month)
        dates, month = self.get_month_data()
        dates = sorted(dates)
        program.write_month(dates, month)
        program.sum_all()
        program.end()

    def get_file_name(self):
        quarter = input("Który kwartał?: ")
        year = input("Wpisz rok: ")
        file_name = f"Kwartał {quarter} {year}.xlsx"
        return file_name, year
    
    def get_month_data(self):
        month = int(input("Podaj numer miesiąca: "))
        dates = []
        while True:
            date = input("Wpisz numer dnia (jeśli chcesz przejśc do nastepnego miesiaca wpisz cos innego niz numer): ")
            if not date.isdigit():
                break
            data = []
            print(f"{date}.{month:02} ZAŁADUNEK")
            print(f"NBP")
            nbp_hundreds = input("Liczba 100: ")
            nbp_fifties = input("Liczba 50: ")
            nbp_twenties = input("Liczba 20: ")
            print(f"WŁASNE")
            hundreds = input("Liczba 100: ")
            fifties = input("Liczba 50: ")
            twenties = input("Liczba 20: ")
            print(f"{date}.{month:02} ROZŁADUNEK")
            print(f"NBP")
            nbp_hundreds_o = input("Liczba 100: ")
            nbp_fifties_o = input("Liczba 50: ")
            nbp_twenties_o = input("Liczba 20: ")
            print(f"WŁASNE")
            hundreds_o = input("Liczba 100: ")
            fifties_o = input("Liczba 50: ")
            twenties_o = input("Liczba 20: ")
            data.append(int(nbp_hundreds))
            data.append(int(nbp_fifties))
            data.append(int(nbp_twenties))
            data.append(int(hundreds))
            data.append(int(fifties))
            data.append(int(twenties))
            data.append(int(nbp_hundreds_o))
            data.append(int(nbp_fifties_o))
            data.append(int(nbp_twenties_o))
            data.append(int(hundreds_o))
            data.append(int(fifties_o))
            data.append(int(twenties_o))
            dates.append((int(date), data))
        return dates, month

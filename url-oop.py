import pandas as pd
import matplotlib.pyplot as plt


class MonthlyDemand(object):
    def __init__(self, region, cust_type):
        self._region = region
        self._cust_type = cust_type
        self.raw_df = None

    def gen_data_url(self, month, year):
        region_mapping = {
            "ภาพรวม กฟภ.": 13,
            "ภาคเหนือ": 14,
            "ภาคตะวันออกเฉียงเหนือ": 15,
            "ภาคกลาง": 16,
            "ภาคใต้": 17,
            "กฟน.1": 1,
            "กฟน.2": 2,
            "กฟน.3": 3,
            "กฟฉ.1": 4,
            "กฟฉ.2": 5,
            "กฟฉ.3": 6,
            "กฟก.1": 7,
            "กฟก.2": 8,
            "กฟก.3": 9,
            "กฟต.1": 10,
            "กฟต.2": 11,
            "กฟต.3": 12,
        }
        cust_type_mapping = {
            "ภาพรวม": 0,
            "บ้านอยู่อาศัย < 150 หน่วย": 10,
            "บ้านอยู่อาศัย > 150 หน่วย": 11,
            "กิจการขนาดเล็ก": 20,
            "กิจการขนาดกลาง": 30,
            "กิจการขนาดใหญ่": 40,
            "กิจการเฉพาะอย่าง": 50,
            "ส่วนราชการฯ": 60,
            "สูบน้ำเพื่อการเกษตร": 70,
        }
        month_mapping = {
            "มกราคม": 1,
            "กุมภาพันธ์": 2,
            "มีนาคม": 3,
            "เมษายน": 4,
            "พฤษภาคม": 5,
            "มิถุนายน": 6,
            "กรกฎาคม": 7,
            "สิงหาคม": 8,
            "กันยายน": 9,
            "ตุลาคม": 10,
            "พฤศจิกายน": 11,
            "ธันวาคม": 12,
        }
        year_mapping = {
            "2555": 12,
            "2556": 13,
            "2557": 14,
            "2558": 15,
            "2559": 16,
            "2560": 17,
            "2561": 18,
            "2562": 19,
            "2563": 20,
            "2564": 21,
            "2565": 22,
            "2566": 23,
        }

        region_code = region_mapping.get(self._region, 0)
        year_code = year_mapping.get(year, 0)
        month_code = month_mapping.get(month, 0)
        cust_type_code = cust_type_mapping.get(self._cust_type, 0)

        base_url = "http://peaoc.pea.co.th/loadprofile/files/%.2d/%s"
        base_fname = "dt%.2d%.2d%.2d%.2d.xls"
        fname = base_fname % (region_code, year_code,
                              month_code, cust_type_code)
        url = base_url % (region_code, fname)

        return url

    def import_data(self, month, year):
        url = self.gen_data_url(month, year)
        self.raw_df = pd.read_excel(url, sheet_name="Source", header=4)

    def __repr__(self):
        return f"{self._region}({self._cust_type}): {self._month}-{self._year}"

    @property
    def _month(self):
        return self.raw_df.columns[1]

    @property
    def _year(self):
        return self.raw_df.columns[0]

    def raw_data(self):
        return self.raw_df


class HourlyDemand(MonthlyDemand):
    def __init__(self, region, cust_type):
        super().__init__(region, cust_type)
        self.hourly_df = None

    def gen_hourly_data(self, day):
        # Convert the 'TIME' column to datetime format with the specified format 'HH:mm'
        self.raw_df['TIME'] = pd.to_datetime(
            self.raw_df['TIME'], format='%H:%M')

        # Set the 'TIME' column's date to match the input day's date
        self.raw_df['TIME'] = self.raw_df['TIME'].apply(
            lambda x: x.replace(year=day.year, month=day.month, day=day.day))

        # Filter the data for the specific day
        day_data = self.raw_df[self.raw_df['TIME'].dt.date ==
                               day.date()].copy()

        # Set the 'TIME' column as the DataFrame index
        day_data.set_index('TIME', inplace=True)

        # Group by hourly frequency and calculate the mean for each hour
        self.hourly_df = day_data.groupby(pd.Grouper(freq='H')).mean()

        # Fill any missing data for the 'HOLIDAY' column with 0
        self.hourly_df['HOLIDAY'].fillna(0, inplace=True)

    def gen_hourly_data(self, day):
        # Convert the 'TIME' column to datetime format with the specified format 'HH:mm'
        # ดังเดิมเราใช้ format='%H:%M' ในการแปลงข้อมูลเวลาในคอลัมน์ 'TIME'
        self.raw_df['TIME'] = pd.to_datetime(
            self.raw_df['TIME'], format='%H:%M', errors='coerce')

    # Handle the special case of "24:00" by converting it to "00:00" of the next day
    # ในกรณีที่เกิด ValueError จากเวลาที่ไม่ถูกต้อง เช่น "24:00" จะถูกแปลงเป็น "00:00" ของวันถัดไป
        self.raw_df.loc[self.raw_df['TIME'] == pd.Timestamp(
            'NaT'), 'TIME'] = pd.to_datetime('00:00') + pd.DateOffset(days=1)

    # Set the 'TIME' column's date to match the input day's date
    # เรากำหนดว่าเวลาในคอลัมน์ 'TIME' จะต้องเป็นวันที่เดียวกับวันที่ input 'day'
        self.raw_df['TIME'] = self.raw_df['TIME'].apply(
            lambda x: x.replace(year=day.year, month=day.month, day=day.day))

    # Filter the data for the specific day
    # คัดกรองข้อมูลใน DataFrame ให้เหลือเฉพาะข้อมูลในวันที่กำหนดในพารามิเตอร์ 'day'
        day_data = self.raw_df[self.raw_df['TIME'].dt.date ==
                               day.date()].copy()

    # Set the 'TIME' column as the DataFrame index
    # กำหนดคอลัมน์ 'TIME' เป็น index ของ DataFrame
        day_data.set_index('TIME', inplace=True)

    # Group by hourly frequency and calculate the mean for each hour
    # จัดกลุ่มข้อมูลตามช่วงเวลาที่เป็นชั่วโมงและคำนวณค่าเฉลี่ยของแต่ละชั่วโมง
        self.hourly_df = day_data.groupby(pd.Grouper(freq='H')).mean()

    # Fill any missing data for the 'HOLIDAY' column with 0
    # กรอกข้อมูลที่หายไปในคอลัมน์ 'HOLIDAY' ด้วยค่า 0
        self.hourly_df['HOLIDAY'].fillna(0, inplace=True)

    def import_data(self, month, year, day):
        super().import_data(month, year)
        self.gen_hourly_data(day)

    def __repr__(self):
        return f"{self._region}({self._cust_type}): {self._month}-{self._year}"


def gen_demand_list(region, cust_type, days):
    demand_list = []
    for day in days:
        # สร้างวัตถุคลาส HourlyDemand โดยใช้ region และ cust_type ที่กำหนด
        hourly_demand = HourlyDemand(region, cust_type)
        # นำเข้าข้อมูลตามวันที่ที่กำหนดในลูป
        hourly_demand.import_data('มกราคม', '2565', pd.to_datetime(day))
        # เพิ่ม DataFrame ในรูปแบบของ hourly_demand.hourly_df เข้าไปในรายการ demand_list
        demand_list.append(hourly_demand.hourly_df)
    return demand_list


# โค้ดทดสอบ
home_demand = MonthlyDemand('ภาพรวม กฟภ.', 'บ้านอยู่อาศัย > 150 หน่วย')
home_demand.import_data('มกราคม', '2565')
print(home_demand)
print(home_demand.raw_data())


home_demand = HourlyDemand('ภาพรวม กฟภ.', 'บ้านอยู่อาศัย > 150 หน่วย')
home_demand.import_data('มกราคม', '2565', pd.to_datetime('2022-01-04'))
print(home_demand)
print(home_demand.hourly_df)

# โค้ดทดสอบ
days = ['2022-01-02', '2022-01-03', '2022-01-04']
demand_list = gen_demand_list('ภาพรวม กฟภ.', 'บ้านอยู่อาศัย > 150 หน่วย', days)
long_df = pd.concat(demand_list)
plt.plot(long_df)
plt.xlabel('Time')
plt.ylabel('Demand (kW)')
plt.show()

# URL generation
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

region = {
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
cust_type = {
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

month = {
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

year = {
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


def gen_data_url(region_input, cust_type_input, month_input, year_input):
    base_url = "http://peaoc.pea.co.th/loadprofile/files/%.2d/%s"
    base_fname = "dt%.2d%.2d%.2d%.2d.xls"

    region_code = region.get(region_input, 0)
    year_code = year.get(year_input, 0)
    month_code = month.get(month_input, 0)
    cust_type_code = cust_type.get(cust_type_input, 0)

    fname = base_fname % (region_code, year_code, month_code, cust_type_code)
    url = base_url % (region_code, fname)

    return url


# Test the function with the provided data
url = gen_data_url(
    'ภาพรวม กฟภ.', 'บ้านอยู่อาศัย > 150 หน่วย', 'มกราคม', '2565')
print(url)


def import_data(url):

    df = pd.read_excel(url, sheet_name="Source", header=4)
    return df


df = import_data(url)
print(df)

# print(df.describe())


# def plot_data(df, day):
#     # สร้างกราฟแสดงแนวโน้มของข้อมูลตามประเภทของวัน
#     plt.figure(figsize=(12, 6))  # กำหนดขนาดของกราฟ

#     # เลือกข้อมูลในประเภทของวันที่ต้องการ
#     data_day = df[day]

#     # สร้างกราฟเส้นแท่งของแต่ละช่วงเวลา
#     time_intervals = df.index
#     plt.plot(time_intervals, data_day, linestyle='-')

#     # ปรับแต่งกราฟ
#     plt.xlabel('Index')
#     plt.ylabel('Demand (KW)')
#     # plt.title(f'Demand Trend on {day}')
#     plt.xticks(rotation=45)
#     plt.grid(True)

#     plt.tight_layout()
#     plt.show()


# # โค้ดทดสอบ
# plot_data(df, 'WORKDAY')


def clean_data(df):
    valid_df = df.copy()
    valid_df.iloc[0:96, 1:] = df.iloc[1:97, 1:]
    valid_df.drop(96, inplace=True)
    valid_df.iloc[95, 1:] = valid_df.iloc[94, 1:]
    # แปลง Index เป็น datetime object
    valid_df.index = pd.to_datetime(valid_df.index)
    return valid_df


valid_df = clean_data(df)
print(valid_df)


def gen_hourly_data(df, day):
    # Convert the 'TIME' column to datetime format with the specified format 'HH:mm'
    df['TIME'] = pd.to_datetime(df['TIME'], format='%H:%M')

    # Set the 'TIME' column's date to match the input day's date
    df['TIME'] = df['TIME'].apply(lambda x: x.replace(
        year=day.year, month=day.month, day=day.day))

    # Filter the data for the specific day
    day_data = df[df['TIME'].dt.date == day.date()].copy()

    # Set the 'TIME' column as the DataFrame index
    day_data.set_index('TIME', inplace=True)

    # Group by hourly frequency and calculate the mean for each hour
    hourly_df = day_data.groupby(pd.Grouper(freq='H')).mean()

    # Fill any missing data for the 'HOLIDAY' column with 0
    hourly_df['HOLIDAY'].fillna(0, inplace=True)

    return hourly_df


def plot_hourly_data(df, day):
    # Convert the input day string to datetime
    day = datetime.strptime(day, '%Y-%m-%d')

    # Generate hourly data for the specified day
    hourly_df = gen_hourly_data(df, day)

    # Plot the data
    plt.figure(figsize=(10, 6))
    plt.plot(hourly_df.index.hour, hourly_df['SATURDAY'], label='Saturday')
    plt.plot(hourly_df.index.hour, hourly_df['SUNDAY'], label='Sunday')
    plt.plot(hourly_df.index.hour, hourly_df['WORKDAY'], label='Workday')

    # Decorate the plot
    plt.xlabel('Time')
    plt.ylabel('Denad (kW)')
    plt.legend()
    plt.grid()
    plt.xticks(range(24))

    # Show the plot
    plt.show()


# Assuming 'valid_df' is the cleaned DataFrame from the previous step.
# You may need to modify this based on your actual DataFrame structure.
valid_df = clean_data(df)

# Test the function with the provided data
plot_hourly_data(valid_df, '2022-01-02')
plot_hourly_data(valid_df, '2022-01-03')
plot_hourly_data(valid_df, '2022-01-04')

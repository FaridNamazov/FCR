def fcr():
    import pandas as pd
    import warnings
    import calendar

    warnings.simplefilter('ignore')

    print("------------- FCR --------------\nClose destination file before start!!!\nEnter month number from 1 to 12")
    while True:
        try:
            month = int(input("For which month are you calculating FCR?: "))
            if month > 12 or month < 1:
                print("There is no such month, month number can vary only from 1 to 12")
                continue
            else:
                break
        except ValueError:
            print("Please enter only whole number from 1 to 12")
            continue
    print("Calculation in progress...")
    fcr = pd.read_excel(r'G:\Customer Service Center\Statistics & Dashboard\FCR\FCR Weekly\FCR_Report.xlsx')
    try:
        fcr2 = pd.read_excel(r'G:\Customer Service Center\Statistics & Dashboard\FCR\FCR Weekly\FCR_Report.xlsx', sheet_name='logs(1)', header=None)
        fcr2.columns = ['MSISDN', 'Date', 'Topic', 'Username', 'Type']
        fcr = pd.concat([fcr, fcr2], ignore_index=True, sort=False)
    except:
        pass

    cal = {month: index for index, month in enumerate(calendar.month_abbr) if month}
    name = ''.join([i for i, m in cal.items() if m == month])
    if name == "Jan":
        name2 = "Dec"
    else:
        name2 = ''.join([i for i, m in cal.items() if m == (month - 1)])

    fcr = fcr.sort_values(['MSISDN', 'Date'], ascending=(True, True))
    fcr.insert(5, "Date_only", pd.to_datetime(fcr['Date']).dt.date)
    fcr.insert(6, "Month", pd.to_datetime(fcr['Date']).dt.month)
    fcr = fcr.drop_duplicates(subset=['MSISDN', 'Username', 'Date_only'])

    fcr = fcr.drop(fcr[fcr.MSISDN.isin([994518179252, 994502210439, 994502310498, 994502312135, 994502312415, 994502310065,
                                        994502312416, 994502312246, 994502312433, 994505049843, 994502312094, 994502312318, 994502312381,
                                        994502310221, 994502502312, 994504502314, 994502312413, 994502312119, 994502312455,
                                        94502312374, 994502312347, 994502311755, 994502312449, 994502311765, 994503502319,
                                        994502312048, 994502312235, 994502210436, 994518179259, 994502100067,994503428108,
                                        994514175583])].index)

    fcr = fcr.drop(fcr[fcr.Topic.isin(['Internal','Complaints checking','TM - Baki ve Gence Elaqe Merkezi qrupunun temsilcisi terefinden Azercell Fintek xidmeti ile bagli zengin cavablandirilmasi (8439)','Corporate - Complaints checking', 'Disconnect from OCS','Internal technical problems', 'Sales call',
                                       '1. Internal use','Fintek: Money transfer ','Fintek Payment (CPM kateqoriyasi)','Fintek Fraud (CPM kateqoriyasi)','Fintek E-manat (CPM kateqoriyasi)','*3443 - Fintek Elaqe Merkezi','Onlayn Musteriye Destek - Azercell Fintek xidmeti ile bagli sorgularin cavablandirilmasi (8438)','Fintek Loyalty (CPM kateqoriyasi)','Fintek: E-manat','Fintek: Fraud','TM - Baki ve Gence Elaqe Merkezi - Unified Solutions Portal portali ile abunecinin melumatlarinin yoxlanilmasi (Azercell Fintek) (8467)','Fintek: Top-up','Fintek Top-up (CPM kateqoriyasi)','Fintek: Chargeback','Fintek: Loyalty','Fintek Chargeback (CPM kateqoriyasi)','Fintek: Onboarding','Fintek: Other','Fintek Onboarding (CPM kateqoriyasi)','Fintek Other (CPM kateqoriyasi)','Fintek: Payment','Fintek Money transfer (CPM kateqoriyasi)', 'Azercell Fintek (Reqemsal kart, Akart)','Survey calls','Welcome Call','Istifade qaydalari (Azercell Fintek)','Fintek','SIMA mobil tetbiqi (Azercell Fintek)','SIMA mobil tetbiqi (Azercell Fintek)','Qosulma qaydasi (Azercell Fintek)','Fintek (CPM kateqoriyasi)','TM - Baki ve Gence Elaqe Merkezi - Unified Solutions Portal portali ile abunecinin balansinin ve odenis tarixcesinin yoxlanilmasi (Azercell Fintek) (8468)','Onlayn cat - abunecinin balansinin ve odenis tarixcesinin yoxlanilmasi (Azercell Fintek) (8472)','SC- Seqmente aid olmayan istifadeci muraciet etdikde (Azercell Fintek)','780092 - Fintech komissiyasiz pul kocurmeleri','Bildirisler (Azercell Fintek)',
'Internal technical problems (CPM kateqoriyasi)'])].index)

    fcr.insert(7, "Apply_count", fcr.groupby(['MSISDN', 'Topic'])['MSISDN'].transform('count'))
    fcr.insert(8, "Repeated", "")
    fcr.loc[fcr['Apply_count'] > 1, ['Repeated']] = 'YES'
    fcr.loc[fcr['Apply_count'] == 1, ['Repeated']] = 'NO'
    fcr = fcr.sort_values(['MSISDN', 'Repeated', 'Topic', 'Date'], ascending=(True, True, True, True))
    fcr.insert(9, "Compare_with", "")
    fcr.loc[(fcr['Repeated'] == "YES") & (fcr['Repeated'].shift(1) == 'NO'), ['Compare_with']] = pd.to_datetime(fcr['Date_only']).dt.date.shift(-1)
    fcr.loc[(fcr['Repeated'] == "YES") & (fcr['Repeated'].shift(1) == 'YES') & (fcr['MSISDN'] == fcr['MSISDN'].shift(1)) & (fcr['Topic'] == fcr['Topic'].shift(1)), ['Compare_with']] = pd.to_datetime(fcr['Date_only']).dt.date.shift(1)
    fcr.loc[(fcr['Repeated'] == "YES") & (fcr['Repeated'].shift(1) == 'YES') & (fcr['MSISDN'] != fcr['MSISDN'].shift(1)), ['Compare_with']] = pd.to_datetime(fcr['Date_only']).dt.date.shift(-1)
    fcr.loc[(fcr['Repeated'] == "YES") & (fcr['Repeated'].shift(1) == 'YES') & (fcr['MSISDN'] == fcr['MSISDN'].shift(1)) & (fcr['Topic'] != fcr['Topic'].shift(1)), ['Compare_with']] = pd.to_datetime(fcr['Date_only']).dt.date.shift(-1)
    fcr.loc[fcr['Repeated'] != "YES", ['Compare_with']] = fcr['Date_only'] - pd.Timedelta(days=+100)
    fcr.insert(10, "Days_difference", (fcr['Date_only'] - fcr['Compare_with']).dt.days)
    fcr.insert(11, "FCR", "")
    fcr.loc[(fcr['Days_difference'] >= -30) & (fcr['Days_difference'] <= 30), ['FCR']] = 'Recall'
    fcr.loc[fcr['FCR'] != "Recall", ['FCR']] = "FCR"

    fcr.insert(12, "FCR_by_agent", "")
    fcr.loc[(fcr['FCR'] == 'Recall') & (fcr['FCR'].shift(-1) == 'Recall') & (fcr['MSISDN'] == fcr['MSISDN'].shift(-1)) & (fcr['Topic'] == fcr['Topic'].shift(-1)), ['FCR_by_agent']] = 'Recall'
    fcr.loc[(fcr['FCR_by_agent'] != 'Recall'), ['FCR_by_agent']] = 'FCR'

    fcr_last_month = fcr.query("Month == @month")

    with pd.ExcelWriter(r'G:\Customer Service Center\Statistics & Dashboard\FCR\FCR Weekly\FCR_Report_ready.xlsx') as writer:
        fcr.to_excel(writer, sheet_name=f'{name2}-{name}', index=False)
        fcr_last_month.to_excel(writer, sheet_name=f'{name}', index=False)

    all_applies = fcr_last_month.shape[0]
    fcr_applies = fcr_last_month[fcr_last_month.FCR == "FCR"].shape[0]

    print('Calculation completed')
    try:
        print(f'Current FCR rate for {name} is: {round((fcr_applies / all_applies) * 100.0, 1)}%')
    except ZeroDivisionError:
        print(f"Can't calculate FCR, no data for month {name}")

fcr()
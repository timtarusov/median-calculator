import pandas as pd
import os
from PyInquirer import prompt, Separator


class CLI:
    script_dir = os.path.dirname(os.path.realpath(__file__))
    questions=[
        {
        "type":"input",
        "name": "empl_data",
        "message": "Введите путь к ШР"
        
    },
        {
        "type":"input",
        "name": "rzp",
        "message": "Введите путь к экспорту из SAP"
    },
        {
        "type":"input",
        "name": "market_data",
        "message": "Введите путь к внешним вилкам"
    },           
        {
        "type":"checkbox",
        "name": "which",
        "choices": [
            Separator("Аналитики"),
            {
            "name": "TTC"
            },
            {
            "name": "median"
            },
            {
            "name": "paymix"
            },
                    ],
        "message": "Выберите нужные таблички"
    }]
    def __init__(self):
        self.answers = prompt(self.questions)

    def calculate_metrics(self):
        mc = MedianCalculator(self.script_dir+self.answers['empl_data'], self.script_dir+self.answers['rzp'])
        mc.group_empls()
        mc.compare_with_market( self.script_dir+self.answers['market_data'])
        mc.make_pivot_for_chart()
        mc.calculate_paymix()
        mc.to_excel(which=self.answers['which'])




class MedianCalculator:
    """Calculate internal median by roles from a given range of employees
    """
    city_code_map = {
        7001: "Москва",
        7002: "Череповец",
        7009: "Колпино",
        7008: "Санкт-Петербург",
        1021: "Череповец"
    }
    
    
    def __init__(self, empl_df_path, rzp_path):
        """Construct a new object for calculating median

        Args:
            empl_df_path (str): path to the xlsx file with roles data. Must have columns "id", "role", "dpts_ext" 
            and "dpt" (optional)
            rzp_path (str): path to the export from SAP
        """

        
        self.employees = pd.read_excel(empl_df_path, engine="openpyxl")
        self.rzp = pd.read_excel(rzp_path, engine="openpyxl")
        self.dirpath = os.path.dirname(empl_df_path)
        
        assert "id" in  self.employees.columns, "id columns must be in df"
        assert "role" in  self.employees.columns, "role columns must be in df"
        assert "dpts_ext" in  self.employees.columns, "dpts_ext columns must be in df"
        
        self.employees = self.employees.merge(self.rzp.loc[:,['Табельный номер', 'З/плата в год']], how="left",
                                              left_on="id", 
                                              right_on="Табельный номер")
        
    def group_empls(self):
        """Group the employees by role and (if possible) by dpt, then calculate the
        median within the group
        """
        MAX_GROUPS = ['dpt', 'role']
        actual_groups = [g for g in MAX_GROUPS if g in self.employees]
        self.grouped = self.employees.groupby(actual_groups)['З/плата в год'].median().reset_index()
        
    def compare_with_market(self, market_data):
        """Compare the pay with the market median
        Args:
            path (str): path to market data file
            
        """
        self.market_data = pd.read_excel(market_data, engine="openpyxl")

        
        self.rzp['dpts_ext'] = self.rzp.merge(self.employees, how="left",left_on="Табельный номер",
                                              right_on="id")['dpts_ext']
        #in case no grade is set, use default = 16
        self.rzp['Раздел персонала_грейд'] = self.rzp['Раздел персонала_грейд'].fillna("_16")
        self.rzp['grade'] = self.rzp['Раздел персонала_грейд'].str.extract(r"_([1-2][1-9])").astype(int)
        
        def map_city_codes(x):
            try:
                return self.city_code_map[x]
            except KeyError:
                return "Череповец"
        
        self.rzp['city'] = self.rzp['РаздПерс'].apply(map_city_codes)
        
        self.rzp['median'] = self.rzp.merge(self.market_data, how="left", 
                                            left_on=['city', 'grade', 'dpts_ext'],
                                            right_on=['city', 'grade', 'dept'])['median']
        
        def place_within_market(x):
            if x['РЗП Месяц'] < x['median']*0.9:
                return "below"
            elif x['РЗП Месяц'] > x['median']*1.1:
                return "above"
            elif (x['РЗП Месяц'] >= x['median']*0.9) and (x['РЗП Месяц'] <= x['median']*1.1):
                return "within"
            else:
                return "unknown"
            
        self.rzp['place_within_market'] = self.rzp.apply(place_within_market, axis=1)
        self.employees['place_within_market'] = self.employees.merge(self.rzp, how='left',
                                                                     left_on="id",
                                                                     right_on="Табельный номер")['place_within_market']
        
        #print(self.rzp.loc[:, ['grade','dpts_ext', 'city','РаздПерс', 'median', 'place_within_market']])
        #print(self.rzp.dtypes)
        #print(self.employees)
    
    def make_pivot_for_chart(self):
        """Calculate the pivot table with comparison of salaries to the market
        """
        assert 'place_within_market' in self.employees
        MAX_GROUPS = ['dpt', 'role']
        actual_groups = [g for g in MAX_GROUPS if g in self.employees]
        self.employees = self.employees.dropna(subset=['place_within_market'])
        self.employees = self.employees.loc[self.employees['place_within_market']!="unknown", :]
        
        self.pivot = self.employees.pivot_table(index=actual_groups, columns=["place_within_market"],
                                                values="id", aggfunc="count", fill_value=0,
                                                margins=True).reset_index()
        for t in ['above', 'below', 'within']:
            self.pivot[t+"_pct"] = self.pivot[t]/self.pivot['All']
        
        self.pivot = self.pivot.loc[:,
                [*actual_groups,'below', 'below_pct', 'within', 'within_pct', 'above', 'above_pct', 'All']]
        
     
    def calculate_paymix(self):
        """Calculate the constant/variable pay ratio
        """
        MAX_GROUPS = ['dpt', 'role']
        actual_groups = [g for g in MAX_GROUPS if g in self.employees]
        self.rzp['paymix'] = self.rzp['2052 Годовая. премия руб.']/self.rzp['З/плата в год']
        self.employees['paymix'] = self.employees.merge(self.rzp, how="left",
                                                        left_on="id", right_on="Табельный номер")['paymix']
        self.paymix = self.employees.groupby(by=actual_groups)['paymix'].mean().reset_index()

           
    def to_excel(self, which=[]):
        """Write the results to excel

        Args:
            path (str): path to excel file
            which (str): whic table to write. 
            -TTC - total target cash for the model
            -median - median distribution vs market
            -paymix - constant/variable ratio
        """
        mapping = {
            "TTC": self.grouped,
            "median": self.pivot,
            "paymix": self.paymix
        }
        for w in which:
            try:
                mapping[w].to_excel(os.path.join(self.dirpath, w+".xlsx"), index=False)
            except KeyError:
                print("Calculate first!")
        

if __name__ == '__main__':
    
    
    # mc = MedianCalculator("./дтрк/ШР_ДТРК.xlsx", "./дтрк/EXPORT_ДТРК.xlsx")
    # mc.group_empls()
    # mc.compare_with_market("./external_median.xlsx")
    # mc.make_pivot_for_chart()
    # mc.calculate_paymix()
    
    # mc.to_excel("./дтрк/internal_median.xlsx", "TTC")
    # mc.to_excel("./дтрк/median.xlsx", "median")
    # mc.to_excel("./дтрк/paymix.xlsx", "paymix")

    
    cli = CLI()
    cli.calculate_metrics()
    # print(os.path.dirname(os.path.realpath(__file__)))
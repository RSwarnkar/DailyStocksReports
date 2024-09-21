# Imports:
import pandas as pd


df_global_debug = pd.DataFrame([])
DEBUG_ME = False


## ============================================================================================================
# Scrape single stock function
def ScrapeSingleStock(stock_name,stock_url,report_folder):
    import pandas as pd 
    from bs4 import BeautifulSoup
    import requests as rq
    import re
    import datetime
    import os
    import time

    # local variables
    stock_sector                = ""
    stock_industry              = ""
    stock_url_type              = ""
    df_basic_info               = pd.DataFrame([]) # Basic Info
    df_top_ratios               = pd.DataFrame([]) # Top Ratios
    df_quaterly_results         = pd.DataFrame([]) # Quarterly Results
    df_profit_n_loss            = pd.DataFrame([]) # Profit & Loss
    df_compounded_sales_growth  = pd.DataFrame([]) # Compounded Sales Growth
    df_compounded_profit_growth = pd.DataFrame([]) # Compounded Profit Growth
    df_stock_price_cagr         = pd.DataFrame([]) # Stock Price CAGR
    df_return_on_equity         = pd.DataFrame([]) # Return on Equity
    df_balance_sheet            = pd.DataFrame([]) # Balance Sheet
    df_cash_flows               = pd.DataFrame([]) # Cash Flows
    df_ratios                   = pd.DataFrame([]) # Ratios
    df_shareholding_pattern     = pd.DataFrame([]) # Shareholding Pattern
    writer                      = None
    report_name                 = ""
    stock_reportpath            = ""
    

    if(DEBUG_ME): print("Info: Starting ScrapeStock()")
    if(DEBUG_ME): print("Info: Stock: ", stock_name)
    if(DEBUG_ME): print("Info: URL: ", stock_url)
    if(DEBUG_ME): print("Info: Report Folder: ", report_folder)

    if not os.path.exists(report_folder):
        os.makedirs(report_folder)
        if(DEBUG_ME): print("Info: Folder created: ",report_folder)


    if 'consolidated' in stock_url:
        stock_url_type = "consolidated"
        if(DEBUG_ME): print("Info: stock_url_type is : ",stock_url_type)
    elif 'consolidated' not in stock_url:
        stock_url_type = "standalone"
        if(DEBUG_ME): print("Info: stock_url_type is : ",stock_url_type)
    else:
        stock_url_type = ""
        if(DEBUG_ME): print("Info: stock_url_type is : ",stock_url_type)


    # Read the page: 
    if(DEBUG_ME): print("Info: Fetching URL...")
    response = rq.get(stock_url)
    if(DEBUG_ME): print("Info: URL Fetch completed")
    if(DEBUG_ME): print("Info: Response status code: ",response.status_code)
    time.sleep(1) # Seconds

    tables = pd.read_html(response.content)
    if(DEBUG_ME): print("Info: Parsing Tables using Pandas. Number of Tables found: ",str(len(tables)))

    if(DEBUG_ME): print("Info: beginning data parsing")

    soup = BeautifulSoup(response.content, "html.parser")
    if(DEBUG_ME): print("Info: BSoup Object loaded")

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Parsing: Sector and industry: Start")
    temp0 = soup.findAll('p', attrs={'class':'sub'})[1].text
    temp0 = ' '.join(temp0.split())

    if m := re.match(r'Sector:\s*([A-Za-z- ]*)Industry:\s*([A-Za-z- ]*)$', temp0):
        stock_sector = m.group(1).strip()
        stock_industry = m.group(2).strip()
        if(DEBUG_ME): print("Info: Parsing: Sector and industry: Found Sector:", stock_sector)
        if(DEBUG_ME): print("Info: Parsing: Sector and industry: Found Industry:", stock_industry)
    else: 
        stock_sector = 'Null'
        stock_industry = 'Null'
        if(DEBUG_ME): print("Error: Parsing: Failed to find Sector and Industry!")
    if(DEBUG_ME): print("Info: Parsing: Sector and industry: Finished")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Parsing: Start - Basic Info")

    temp_1 = [{
    "Stockname":stock_name,
    "Stock Report Type": stock_url_type,
    "Stock Url": stock_url,
    "Sector": stock_sector,
    "Industry": stock_industry
    }]

    df_basic_info = pd.DataFrame(temp_1)

    if(DEBUG_ME): print("Info: Parsing: Finished - Basic Info")
    ## ---------------------------------------------------------

    


    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Parsing: Start - Top Ratios")

    ul = soup.find("ul#top-ratios")  # Selector of top ratios
    lines = []
    for ul in soup.findAll('ul', id='top-ratios'):
        for li in ul.findAll('li'):
            li_parsed_text = li.text
            li_parsed_text = re.sub('[\s ]+', ' ', li_parsed_text)
            li_parsed_text = li_parsed_text.strip()
            #print(li_parsed_text)
            lines.append(li_parsed_text)

    if(DEBUG_ME): print("Info: Parsing: Top Ratios: Cleaning up the To Ratios data line by line")
     # Line - 0
    if m := re.match(r"Market Cap ₹ ([0-9,.-]+) Cr.", lines[0]):
        line_0 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found Market Cap: ", line_0)
    else: 
        line_0 = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found Market Cap")

    # Line - 1
    if m := re.match(r"Current Price ₹ ([0-9,.-]+)", lines[1]):
        line_1 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found Current Price: ", line_1)
    else: 
        line_1 = "NaN"

    # Line - 2a
    if m := re.match(r"High \/ Low ₹ ([0-9,..-]+) \/ ([0-9,.]+)", lines[2]):
        line_2a = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found High: ", line_2a)
    else: 
        line_2a = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found High")

    # Line - 2b
    if m := re.match(r"High \/ Low ₹ ([0-9,.-]+) \/ ([0-9,.]+)", lines[2]):
        line_2b = m.group(2).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found Low: ", line_2b)
    else: 
        line_2b = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found Low")

    # Line - 3
    if m := re.match(r"Stock P\/E ([0-9,.-]+)", lines[3]):
        line_3 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found Low: ", line_3)
    else: 
        line_3 = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found __")

    # Line - 4
    if m := re.match(r"Book Value ₹ ([0-9,.-]+)", lines[4]):
        line_4 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found Book Value: ", line_4)
    else: 
        line_4 = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found Book Value")

    # Line - 5
    if m := re.match(r"Dividend Yield ([0-9,.-]+) %", lines[5]):
        line_5 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found Dividend Yield: ", line_5)
    else: 
        line_5 = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found Dividend Yield")

    # Line - 6
    if m := re.match(r"ROCE ([0-9,.-]+) %", lines[6]):
        line_6 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found ROCE: ", line_6)
    else: 
        line_6 = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found ROCE")

    # Line - 7
    if m := re.match(r"ROE ([0-9,.-]+) %", lines[7]):
        line_7 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found ROE: ", line_7)
    else: 
        line_7 = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found ROE")

    # Line - 8
    if m := re.match(r"Face Value ₹ ([0-9,.-]+)", lines[8]):
        line_8 = m.group(1).replace(',','')
        if(DEBUG_ME): print("Info: Parsing: Top Ratios: Found Face Value: ", line_8)
    else: 
        line_8 = "NaN"
        if(DEBUG_ME): print("Error: Parsing: Top Ratios: Not Found Face Value")

    if(DEBUG_ME): print("Info: Parsing: Top Ratios: Preparing Dataframe")
    temp_2 = [{
    "Stockname":stock_name,
    "Market Cap in Cores Rupees":line_0,
    "Current Price in Rupees":line_1,
    "High in Rupees":line_2a,
    "Low in Rupees":line_2b,
    "Stock PE":line_3,
    "Book Value in Rupees":line_4,
    "Dividend Yield %": line_5,
    "ROCE %":line_6,
    "ROE %":line_7,
    "Face Value in Rupees":line_8,
    }]

    df_top_ratios = pd.DataFrame(temp_2)
    if(DEBUG_ME): print("Info: Parsing: Top Ratios: Done Dataframe for Top Ratios")

    if(DEBUG_ME): print("Info: Parsing: Finished - Top Ratios")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Parsing: Start - All Pandas Tables")
    
    df_quaterly_results         = tables[0] # Quarterly Results
    df_profit_n_loss            = tables[1] # Profit & Loss
    df_compounded_sales_growth  = tables[2] # Compounded Sales Growth
    df_compounded_profit_growth = tables[3] # Compounded Profit Growth
    df_stock_price_cagr         = tables[4] # Stock Price CAGR
    df_return_on_equity         = tables[5] # Return on Equity
    df_balance_sheet            = tables[6] # Balance Sheet
    df_cash_flows               = tables[7] # Cash Flows
    df_ratios                   = tables[8] # Ratios
    df_shareholding_pattern     = tables[9] # Shareholding Pattern

    if(DEBUG_ME): print("Info: Parsing: Finished - All Pandas Tables")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Cleaning: Start - All Pandas Tables")
    
    # Cleanup table: Quarterly Results
    df_quaterly_results.rename(columns={'Unnamed: 0':'Quarterly Results'}, inplace=True)
    df_quaterly_results.replace(u"\u00A0\+", "", regex=True,inplace=True)
    if(DEBUG_ME): print("Info: Cleaning: Done - Quarterly Results")
    
    # Cleanup table: Profit & Loss
    df_profit_n_loss.rename(columns={'Unnamed: 0':'Profit and Loss'}, inplace=True)
    df_profit_n_loss.replace(u"\u00A0\+", "", regex=True,inplace=True)
    if(DEBUG_ME): print("Info: Cleaning: Done - Profit & Loss")

    # Cleanup table: Compounded Sales Growth
    df_compounded_sales_growth.replace(":", "", regex=True,inplace=True) 
    
    # Cleanup table: Compounded Profit Growth
    df_compounded_profit_growth.replace(":", "", regex=True,inplace=True) 
    
    # Cleanup table: Stock Price CAGR
    df_stock_price_cagr.replace(":", "", regex=True,inplace=True) 
    
    # Cleanup table: Return on Equity
    df_return_on_equity.replace(":", "", regex=True,inplace=True) 
    
    # Cleanup table: Balance Sheet
    df_balance_sheet.rename(columns={'Unnamed: 0':'Balance Sheet'}, inplace=True)
    df_balance_sheet.replace(u"\u00A0\+", "", regex=True,inplace=True) 
    
    # Cleanup table: Cash Flows
    df_cash_flows.rename(columns={'Unnamed: 0':'Cash Flows'}, inplace=True)
    df_cash_flows.replace(u"\u00A0\+", "", regex=True,inplace=True) 
    
    # Cleanup table: Ratios
    df_ratios.rename(columns={'Unnamed: 0':'Ratios'}, inplace=True)
    
    # Cleanup table: Shareholding Pattern
    df_shareholding_pattern.rename(columns={'Unnamed: 0':'Shareholding Pattern'}, inplace=True)
    df_shareholding_pattern.replace(u"\u00A0\+", "", regex=True,inplace=True)

    if(DEBUG_ME): print("Info: Cleaning: Finished - All Pandas Tables")
    ## ---------------------------------------------------------

    ## =================================================================================
    ## Dataframe contains space, non-numeric and percent sign so those should be removed
    ## =================================================================================
    
    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_top_ratios")
    # First col is texual, rest are numeric
    #df_top_ratios[df_top_ratios.columns[1:]] = df_top_ratios[df_top_ratios.columns[1:]].apply(pd.to_numeric)

    for col in df_top_ratios.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_top_ratios[col] = df_top_ratios[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_top_ratios[col] = df_top_ratios[col].apply(pd.to_numeric)
    
    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_top_ratios.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_top_ratios")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_quaterly_results")
    # First col is texual, rest are numeric
    for col in df_quaterly_results.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_quaterly_results[col] = df_quaterly_results[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_global_debug[col] = df_quaterly_results[col]
        df_quaterly_results[col] = df_quaterly_results[col].apply(pd.to_numeric)
    
    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_quaterly_results.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_quaterly_results")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_profit_n_loss")
    # First col is texual, rest are numeric
    for col in df_profit_n_loss.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_profit_n_loss[col] = df_profit_n_loss[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_profit_n_loss[col] = df_profit_n_loss[col].apply(pd.to_numeric)
    
    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_profit_n_loss.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_profit_n_loss")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_compounded_sales_growth")
    # First col is texual, rest are numeric
    for col in df_compounded_sales_growth.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_compounded_sales_growth[col] = df_compounded_sales_growth[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_compounded_sales_growth[col] = df_compounded_sales_growth[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_compounded_sales_growth.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_compounded_sales_growth")
    ## ---------------------------------------------------------

    
    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_compounded_profit_growth")
    # First col is texual, rest are numeric
    for col in df_compounded_profit_growth.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_compounded_profit_growth[col] = df_compounded_profit_growth[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_compounded_profit_growth[col] = df_compounded_profit_growth[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_compounded_profit_growth.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_compounded_profit_growth")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_stock_price_cagr")
    # First col is texual, rest are numeric
    for col in df_stock_price_cagr.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_stock_price_cagr[col] = df_stock_price_cagr[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_stock_price_cagr[col] = df_stock_price_cagr[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_stock_price_cagr.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_stock_price_cagr")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_return_on_equity")
    # First col is texual, rest are numeric
    for col in df_return_on_equity.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_return_on_equity[col] = df_return_on_equity[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_return_on_equity[col] = df_return_on_equity[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_return_on_equity.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_return_on_equity")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_balance_sheet")
    # First col is texual, rest are numeric
    for col in df_balance_sheet.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        ## CAUTION: Converting value to string!
        df_balance_sheet[col] = df_balance_sheet[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_balance_sheet[col] = df_balance_sheet[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_balance_sheet.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_balance_sheet")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_cash_flows")
    # First col is texual, rest are numeric
    for col in df_cash_flows.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_cash_flows[col] = df_cash_flows[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_cash_flows[col] = df_cash_flows[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_cash_flows.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_cash_flows")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_ratios")
    # First col is texual, rest are numeric
    for col in df_ratios.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_ratios[col] = df_ratios[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_ratios[col] = df_ratios[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_ratios.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_ratios")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Data Transform: Start - df_shareholding_pattern")
    # First col is texual, rest are numeric
    for col in df_shareholding_pattern.columns[1:]:
        # Remove comma, percent sign and Empty value with NaN
        df_shareholding_pattern[col] = df_shareholding_pattern[col].astype(str).str.replace(',','').str.rstrip('%').replace('nan','NaN').replace('','NaN').replace("NaN", pd.NA)
        df_shareholding_pattern[col] = df_shareholding_pattern[col].apply(pd.to_numeric)

    if(DEBUG_ME): print("Info: Data Transform: Outcome: \n",df_shareholding_pattern.info())
    if(DEBUG_ME): print("Info: Data Transform: Done - df_shareholding_pattern")
    ## ---------------------------------------------------------

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Saving Report...")

    sheet_names = ["Basic Info", "Top Ratios","Quarterly Results", "Profit & Loss", "Compounded Sales Growth", "Compounded Profit Growth", 
                   "Stock Price CAGR", "Return on Equity", "Balance Sheet", "Cash Flows", "Ratios", "Shareholding Pattern"]
    dataframes  = [df_basic_info, df_top_ratios, df_quaterly_results, df_profit_n_loss , df_compounded_sales_growth, df_compounded_profit_growth, 
                   df_stock_price_cagr, df_return_on_equity, df_balance_sheet, df_cash_flows, df_ratios, df_shareholding_pattern]

    if 'consolidated' in stock_url:
        report_name = stock_name + "-consolidated-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
        if(DEBUG_ME): print("Info: report name will be: ",report_name)
    elif 'consolidated' not in stock_url:
        report_name = stock_name + "-standalone-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
        if(DEBUG_ME): print("Info: report name suffix will be: ",report_name)
    else:
        report_name = stock_name + "-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
        if(DEBUG_ME): print("Info: report name suffix will be: ",report_name)
    
    
    stock_reportpath = report_folder + "/" + report_name
    
    writer = pd.ExcelWriter(stock_reportpath, engine='xlsxwriter')
    
    for i, frame in enumerate(dataframes):
        frame.to_excel(writer, sheet_name = sheet_names[i], index=False)
    writer.close()
    writer.handles = None

    if(DEBUG_ME): print("Info: Report saved successfully: ", stock_reportpath)
    ## ---------------------------------------------------------

    
    return stock_reportpath
## ============================================================================================================
# Visualize single stock function
def VisualizeSingleStock(stock_excel_filepath,output_folder):
    import pandas as pd 
    from bs4 import BeautifulSoup
    import requests as rq
    import re
    import datetime
    import os
    import time

    import plotly.offline as pyo 
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots

    # local variables
    all_figures = []
    fig01 = go.Figure()
    fig02 = go.Figure()
    fig03 = go.Figure()
    fig04 = go.Figure()
    fig05 = go.Figure()
    fig06 = go.Figure()
    fig07 = go.Figure()
    fig08 = go.Figure()
    fig09 = go.Figure()
    fig10 = go.Figure()
    fig11 = go.Figure()
    fig12 = go.Figure()
    fig13 = go.Figure()
    fig14 = go.Figure()
    fig15 = go.Figure()

    # https://colorkit.co/palettes/8-colors/
    color_set = {
        "dark_red"       : "#6f1926",
        "red"            : "#de324c",
        "orange"         : "#f4895f",
        "yellow"         : "#f8e16f",
        "bright_green"   : "#8bc34a",
        "dark_green"     : "#009688", #"#318a01",
        "blue"           : "#369acc",
        "sky_blue"       : "#80d9ff",
        "dark_blue"      : "#0057e5",
        "violet"         : "#9656a2",
        "purple"         : "#7b4fff",
        "black"          : "#212020",
        "gray"           : "#777675"
    }
    
    df_top_ratios               = pd.DataFrame([]) # Top Ratios
    df_quaterly_results         = pd.DataFrame([]) # Quarterly Results
    df_profit_n_loss            = pd.DataFrame([]) # Profit & Loss
    df_compounded_sales_growth  = pd.DataFrame([]) # Compounded Sales Growth
    df_compounded_profit_growth = pd.DataFrame([]) # Compounded Profit Growth
    df_stock_price_cagr         = pd.DataFrame([]) # Stock Price CAGR
    df_return_on_equity         = pd.DataFrame([]) # Return on Equity
    df_balance_sheet            = pd.DataFrame([]) # Balance Sheet
    df_cash_flows               = pd.DataFrame([]) # Cash Flows
    df_ratios                   = pd.DataFrame([]) # Ratios
    df_shareholding_pattern     = pd.DataFrame([]) # Shareholding Pattern

    stock_name                  = ""
    stock_url                   = ""
    stock_url_type              = ""
    stock_sector                = ""
    stock_industry              = ""
    stock_marketcap             = ""
    
    writer                      = None
    html_report_name            = ""
    stock_html_reportpath       = ""
    

    if(DEBUG_ME): print("Info: Starting VisualizeStock()")
    if(DEBUG_ME): print("Info: Stock: ", stock_name)
    if(DEBUG_ME): print("Info: URL: ", stock_url)
    if(DEBUG_ME): print("Info: Excel File of the stock to load : ", stock_excel_filepath)

    # Load Excel File
    if(DEBUG_ME): print("Info: Loading Excel File")
    xls = pd.ExcelFile(stock_excel_filepath)

    df_basic_info               = pd.read_excel(xls, "Basic Info") # Top Ratios
    df_top_ratios               = pd.read_excel(xls, "Top Ratios") # Top Ratios
    df_quaterly_results         = pd.read_excel(xls, "Quarterly Results") # Quarterly Results
    df_profit_n_loss            = pd.read_excel(xls, "Profit & Loss") # Profit & Loss
    df_compounded_sales_growth  = pd.read_excel(xls, "Compounded Sales Growth") # Compounded Sales Growth
    df_compounded_profit_growth = pd.read_excel(xls, "Compounded Profit Growth") # Compounded Profit Growth
    df_stock_price_cagr         = pd.read_excel(xls, "Stock Price CAGR") # Stock Price CAGR
    df_return_on_equity         = pd.read_excel(xls, "Return on Equity") # Return on Equity
    df_balance_sheet            = pd.read_excel(xls, "Balance Sheet") # Balance Sheet
    df_cash_flows               = pd.read_excel(xls, "Cash Flows") # Cash Flows
    df_ratios                   = pd.read_excel(xls, "Ratios") # Ratios
    df_shareholding_pattern     = pd.read_excel(xls, "Shareholding Pattern") # Shareholding Pattern

    xls.close()

    #df_global_debug = df_quaterly_results.copy()

    if(DEBUG_ME): print(df_quaterly_results.info())

    if(DEBUG_ME): print("Info: Loaded all dataframes")

    ## ---------------------------------------------------------
    if(DEBUG_ME): print("Info: Reading basic info")
    stock_name = df_basic_info.iloc[0]['Stockname']
    stock_url  = df_basic_info.iloc[0]['Stock Url']
    stock_url_type = df_basic_info.iloc[0]['Stock Report Type']
    stock_sector = df_basic_info.iloc[0]['Sector']
    stock_industry = df_basic_info.iloc[0]['Industry']
    stock_marketcap = df_top_ratios.iloc[0]['Market Cap in Cores Rupees']

    if(DEBUG_ME): print("Info: stock_name: ",stock_name)
    if(DEBUG_ME): print("Info: stock_url: ",stock_url)
    if(DEBUG_ME): print("Info: stock_url_type: ",stock_url_type)
    if(DEBUG_ME): print("Info: stock_sector: ",stock_sector)
    if(DEBUG_ME): print("Info: stock_industry: ",stock_industry)
    if(DEBUG_ME): print("Info: stock_marketcap: ",stock_marketcap)
    if(DEBUG_ME): print("Info: Done basic info")
    ## ---------------------------------------------------------

    ## ==================================================================================================================
    ## Chart-01: Fig01 - [Quaterly Results] - `Sales, Expenses, Operating Profit, and Net Profit Trend (Quarterly)`
    ## ==================================================================================================================

    df_quaterly_results = df_quaterly_results.rename(columns={'Quarterly Results': 'Quarter'}) # Rename a col
    df_quaterly_results = df_quaterly_results.set_index('Quarter').T # Transpose with reset index

    plot_cols   = [ 'Sales', 'Revenue', 'Expenses', 'Operating Profit', 'Financing Profit', 'Net Profit']
    plot_colors = {
        'Sales' : color_set.get("dark_blue"),
        'Revenue' : color_set.get("dark_blue"),
        'Expenses' : color_set.get("red"),
        'Operating Profit' : color_set.get("dark_green"),
        'Financing Profit' : color_set.get("dark_green"),
        'Net Profit' : color_set.get("bright_green")
    }

    # Add Traces
    for col in plot_cols:
        if col in df_quaterly_results.columns:
            x = df_quaterly_results.index
            y = df_quaterly_results[col]
            fig01.add_trace(go.Bar(x=x, y=y, name=col, marker_color=plot_colors.get(col))) 
    
    plot_title = '<b>Quarterly Trend: Sales, Expenses, Operating Profit, Net Profit: </b>' + stock_name + " (" + stock_url_type + ")"

    fig01.update_layout( title=plot_title, xaxis_title='Quarter', yaxis_title='Rupees in Cr.')
    updatemenus=[
        dict(
            type="buttons",
            buttons=list([
                dict(
                    label="Group",
                    method="relayout",
                    args=[{"barmode": "group"}] ),
                dict(
                    label="Stack",
                    method="relayout",
                    args=[{"barmode": "stack"}] )
             ]) )]
    fig01.update_layout(updatemenus=updatemenus)
    #fig01.show()
    all_figures.append(fig01)
    if(DEBUG_ME): print("Info: Chart-01 Generated")

    ## ==================================================================================================================
    ## Chart-02: Fig02 - [Quaterly Results] - `Operating or Financing Profit Margin % and Net Profit Margin % Trend (Quarterly)`
    ## ==================================================================================================================
    
    plot_cols   = [ 'Financing Margin %', 'OPM %', 'Net Profit Margin %']
    plot_colors = {
        'Financing Margin %' : color_set.get("dark_green"),
        'OPM %' : color_set.get("dark_green"),
        'Net Profit Margin %' : color_set.get("bright_green")
    }
    plot_title = "Null"

    if 'Sales' in df_quaterly_results.columns:
        y1 = df_quaterly_results['Sales']
        y2 = df_quaterly_results['Net Profit'] 
        df_quaterly_results['Net Profit Margin %'] =  df_quaterly_results['Net Profit'] / df_quaterly_results['Sales']  * 100
        plot_title = '<b>Quarterly Trend: Operating & Net Profit Margin % : </b>' + stock_name + " (" + stock_url_type + ")"

    elif 'Revenue' in df_quaterly_results.columns:
        y1 = df_quaterly_results['Revenue']
        y2 = df_quaterly_results['Net Profit'] 
        df_quaterly_results['Net Profit Margin %'] = df_quaterly_results['Net Profit'] / df_quaterly_results['Revenue'] * 100
        plot_title = '<b>Quarterly Trend: Operating & Net Profit Margin % : </b>' + stock_name + " (" + stock_url_type + ")"
    else:
        df_quaterly_results['Net Profit Margin %'] = 0.0
        plot_title = '<b>Quarterly Trend: OPFM Margin % and Net Profit Margin % : </b>' + stock_name + " (" + stock_url_type + ")"
    
    y = df_quaterly_results['Net Profit Margin %']

    # Add Traces
    for col in plot_cols:
        if col in df_quaterly_results.columns:
            x = df_quaterly_results.index
            y = df_quaterly_results[col]
            fig02.add_trace(go.Bar(x=x, y=y, name=col, marker_color=plot_colors.get(col))) 

    fig02.update_layout( title=plot_title, xaxis_title='Quarter', yaxis_title=' % ')
    updatemenus=[
        dict(
            type="buttons",
            buttons=list([
                dict(
                    label="Group",
                    method="relayout",
                    args=[{"barmode": "group"}] ),
                dict(
                    label="Stack",
                    method="relayout",
                    args=[{"barmode": "stack"}] )
             ]) )]
    fig02.update_layout(updatemenus=updatemenus)
    #fig02.show()
    all_figures.append(fig02)
    if(DEBUG_ME): print("Info: Chart-02 Generated")

    ## ==================================================================================================================
    ## Chart-03: Fig03 - [Quaterly Results] - `EPS Trend (Quarterly)`
    ## ==================================================================================================================

    df_quaterly_results['EPS Change %'] = df_quaterly_results['EPS in Rs'].pct_change()*100

    x  = df_quaterly_results.index
    y1 = df_quaterly_results['EPS in Rs']
    y2 = df_quaterly_results['EPS Change %'].apply(lambda x: round(x, 2))
    
    fig03.add_trace(go.Bar(x=x, y=y1, name='EPS in Rs', marker_color=color_set.get("dark_green"),
                           text=y2,
                           textposition="outside",
                           texttemplate="%{text} %, [%{y}]"
                          ))
    plot_title = '<b>Quarterly Trend: EPS in Rupees : </b>' + stock_name + " (" + stock_url_type + ")"
    fig03.update_layout(title=plot_title, xaxis_title='Quarter', yaxis_title=' Rupees ')
    #fig03.show()
    all_figures.append(fig03)
    if(DEBUG_ME): print("Info: Chart-03 Generated")

    ## ==================================================================================================================
    ## Chart-04: Fig04 - [Quaterly Results] - `NPA Trend (Quarterly)`
    ## ==================================================================================================================

    # This chart is only applicable to Banks and NBFC companies
    if 'Gross NPA %' in df_quaterly_results.columns:
        x  = df_quaterly_results.index
        y1 = df_quaterly_results['Gross NPA %']
        y2 = df_quaterly_results['Net NPA %']
        fig04.add_trace(go.Bar(x=x, y=y1, name='Gross NPA %', marker_color=color_set.get("dark_red")))
        fig04.add_trace(go.Bar(x=x, y=y2, name='Net NPA %', marker_color=color_set.get("red")))
        plot_title = '<b>Quarterly Trend: NPA % : </b>' + stock_name + " (" + stock_url_type + ")"
        fig04.update_layout( title=plot_title, xaxis_title='Quarter', yaxis_title='%')
        updatemenus=[
           dict(
                type="buttons",
                buttons=list([
                    dict(label="Group", method="relayout", args=[{"barmode": "group"}]),
                    dict(label="Stack", method="relayout", args=[{"barmode": "stack"}])
                    ])
            )]
        fig04.update_layout(updatemenus=updatemenus)
        #fig04.show()
        all_figures.append(fig04)
        if(DEBUG_ME): print("Info: Chart-04 Generated")

    # ## ==================================================================================================================
    # ## Chart-05: Fig05 - [Profit and Loss] - `Sales, Expenses, Operating Profit, and Net Profit Trend (Yearly)`
    # ## ==================================================================================================================

    df_profit_n_loss = df_profit_n_loss.rename(columns={'Profit and Loss': 'Year'}) # Rename a col
    df_profit_n_loss = df_profit_n_loss.set_index('Year').T # Transpose with reset index

    plot_cols   = [ 'Sales', 'Revenue', 'Expenses', 'Operating Profit', 'Financing Profit', 'Net Profit']
    plot_colors = {
        'Sales' : color_set.get("dark_blue"),
        'Revenue' : color_set.get("dark_blue"),
        'Expenses' : color_set.get("red"),
        'Operating Profit' : color_set.get("dark_green"),
        'Financing Profit' : color_set.get("dark_green"),
        'Net Profit' : color_set.get("bright_green")
    }

    # Add Traces
    for col in plot_cols:
        if col in df_profit_n_loss.columns:
            x = df_profit_n_loss.index
            y = df_profit_n_loss[col]
            fig05.add_trace(go.Bar(x=x, y=y, name=col, marker_color=plot_colors.get(col))) 
    
    plot_title = '<b>Yearly Trend: Sales, Expenses, Operating Profit, Net Profit: </b>' + stock_name + " (" + stock_url_type + ")"

    fig05.update_layout( title=plot_title, xaxis_title='Quarter', yaxis_title='Rupees in Cr.')
    updatemenus=[
        dict(
            type="buttons",
            buttons=list([
                dict(
                    label="Group",
                    method="relayout",
                    args=[{"barmode": "group"}] ),
                dict(
                    label="Stack",
                    method="relayout",
                    args=[{"barmode": "stack"}] )
             ]) )]
    fig05.update_layout(updatemenus=updatemenus)
    #fig05.show()
    all_figures.append(fig05)
    if(DEBUG_ME): print("Info: Chart-05 Generated")


    ## ==================================================================================================================
    ## Chart-06: Fig06 - [Profit and Loss] - `Operating or Financing Profit Margin % and Net Profit Margin % Trend (Yearly)`
    ## ==================================================================================================================

    plot_cols   = [ 'Financing Margin %', 'OPM %', 'Net Profit Margin %']
    plot_colors = {
        'Financing Margin %' : color_set.get("dark_green"),
        'OPM %' : color_set.get("dark_green"),
        'Net Profit Margin %' : color_set.get("bright_green")
    }
    plot_title = "Null"

    if 'Sales' in df_profit_n_loss.columns:
        y1 = df_profit_n_loss['Sales']
        y2 = df_profit_n_loss['Net Profit'] 
        df_profit_n_loss['Net Profit Margin %'] =  df_profit_n_loss['Net Profit'] / df_profit_n_loss['Sales']  * 100
        plot_title = '<b>Yearly Trend: Operating & Net Profit Margin % : </b>' + stock_name + " (" + stock_url_type + ")"

    elif 'Revenue' in df_profit_n_loss.columns:
        y1 = df_profit_n_loss['Revenue']
        y2 = df_profit_n_loss['Net Profit'] 
        df_profit_n_loss['Net Profit Margin %'] = df_profit_n_loss['Net Profit'] / df_profit_n_loss['Revenue'] * 100
        plot_title = '<b>Yealy Trend: Operating & Net Profit Margin % : </b>' + stock_name + " (" + stock_url_type + ")"
    else:
        df_profit_n_loss['Net Profit Margin %'] = 0.0
        plot_title = '<b>Yearly Trend: OPFM Margin % and Net Profit Margin % : </b>' + stock_name + " (" + stock_url_type + ")"
    
    y = df_profit_n_loss['Net Profit Margin %']

    # Add Traces
    for col in plot_cols:
        if col in df_profit_n_loss.columns:
            x = df_profit_n_loss.index
            y = df_profit_n_loss[col]
            fig06.add_trace(go.Bar(x=x, y=y, name=col, marker_color=plot_colors.get(col))) 

    fig06.update_layout( title=plot_title, xaxis_title='Quarter', yaxis_title=' % ')
    updatemenus=[
        dict(
            type="buttons",
            buttons=list([
                dict(
                    label="Group",
                    method="relayout",
                    args=[{"barmode": "group"}] ),
                dict(
                    label="Stack",
                    method="relayout",
                    args=[{"barmode": "stack"}] )
             ]) )]
    fig06.update_layout(updatemenus=updatemenus)
    #fig06.show()
    all_figures.append(fig06)
    if(DEBUG_ME): print("Info: Chart-06 Generated")

    ## ==================================================================================================================
    ## Chart-07: Fig07 - [Profit and Loss] - `EPS Trend (Yearly)`
    ## ==================================================================================================================

    df_profit_n_loss['EPS Change %'] = df_profit_n_loss['EPS in Rs'].pct_change()*100

    x  = df_profit_n_loss.index
    y1 = df_profit_n_loss['EPS in Rs']
    y2 = df_profit_n_loss['EPS Change %'].apply(lambda x: round(x, 2))
    
    fig07.add_trace(go.Bar(x=x, y=y1, name='EPS in Rs', marker_color=color_set.get("dark_green"),
                           text=y2,
                           textposition="outside",
                           texttemplate="%{text} %, [%{y}]"
                          ))
    plot_title = '<b>Yearly Trend: EPS in Rupees : </b>' + stock_name + " (" + stock_url_type + ")"
    fig07.update_layout(title=plot_title, xaxis_title='Year', yaxis_title=' Rupees ')
    #fig07.show()
    all_figures.append(fig07)
    if(DEBUG_ME): print("Info: Chart-07 Generated")

    ## ==================================================================================================================
    ## Chart-08: Fig08 - [Profit and Loss] - `Dividend Payout % Trend (Yearly)`
    ## ==================================================================================================================

    x  = df_profit_n_loss.index
    y1 = df_profit_n_loss['Dividend Payout %']
        
    fig08.add_trace(go.Bar(x=x, y=y1, name='Dividend Payout %', marker_color=color_set.get("bright_green")))
    plot_title = '<b>Yearly Trend: Dividend Payout % : </b>' + stock_name + " (" + stock_url_type + ")"
    fig08.update_layout(title=plot_title, xaxis_title='Year', yaxis_title=' % ')
    #fig08.show()
    all_figures.append(fig08)
    if(DEBUG_ME): print("Info: Chart-08 Generated")


    ## ==================================================================================================================
    ## Chart-09: Fig09 - [Compounded Sales Growth, Compounded Profit Growth, Stock Price CAGR, Return on Equity] - `Key Fundamental Summary`
    ## ==================================================================================================================

    x1 = df_compounded_sales_growth['Compounded Sales Growth']
    y1 = df_compounded_sales_growth['Compounded Sales Growth.1']

    x2 = df_compounded_profit_growth['Compounded Profit Growth']
    y2 = df_compounded_profit_growth['Compounded Profit Growth.1']

    x3 = df_stock_price_cagr['Stock Price CAGR']
    y3 = df_stock_price_cagr['Stock Price CAGR.1']

    x4 = df_return_on_equity['Return on Equity']
    y4 = df_return_on_equity['Return on Equity.1']

    fig09 = make_subplots(
            rows=2, cols=2,
            specs=[
                [{"type": "bar"}, {"type": "bar"}],
                [{"type": "bar"}, {"type": "bar"}],
            ],
            subplot_titles=('Compounded Sales Growth %', 'Compounded Profit Growth %', 
                            'Stock Price CAGR %', 'Return on Equity %') )

    fig09.add_trace( go.Bar(x=x1, y=y1, name='Compounded Sales Growth %',  marker_color=color_set.get("dark_blue")), row=1, col=1)
    fig09.add_trace( go.Bar(x=x2, y=y2, name='Compounded Profit Growth %', marker_color=color_set.get("bright_green")), row=1, col=2)
    fig09.add_trace( go.Bar(x=x3, y=y3, name='Stock Price CAGR %', marker_color=color_set.get("red")), row=2, col=1)
    fig09.add_trace( go.Bar(x=x4, y=y4, name='Return on Equity %', marker_color=color_set.get("dark_green")), row=2, col=2)

    plot_title = '<b>Yearly Trend: Key Fundamentals : </b>' + stock_name + " (" + stock_url_type + ")"
    fig09.update_layout(title=plot_title)
    #fig09.show()
    all_figures.append(fig09)
    if(DEBUG_ME): print("Info: Chart-09 Generated")

    ## ==================================================================================================================
    ## Chart-10: Fig10 - [Balance Sheet] - `Balance Sheet Trends`
    ## ==================================================================================================================

    df_balance_sheet = df_balance_sheet.rename(columns={'Balance Sheet': 'Year'}) # Rename a col
    df_balance_sheet = df_balance_sheet.set_index('Year').T # Transpose with reset index


    plot_cols = ['Equity Capital', 'Reserves', 'Borrowings', 'Other Liabilities',
                 'Total Liabilities', 'Fixed Assets', 'CWIP', 'Investments',
                 'Other Assets', 'Total Assets']
    plot_colors = {
       'Equity Capital'   : color_set.get("blue"), 
       'Reserves'         : color_set.get("bright_green"),
       'Borrowings'       : color_set.get("red"),
       'Other Liabilities': color_set.get("orange"),
       'Total Liabilities': color_set.get("dark_red"),
       'Fixed Assets'     : color_set.get("violet"),
       'CWIP'             : color_set.get("gray"),
       'Investments'      : color_set.get("sky_blue"),
       'Other Assets'     : color_set.get("purple"),
       'Total Assets'     : color_set.get("dark_blue"),
    }
    plot_title = '<b>Balance Sheet Trends: </b>' + stock_name + " (" + stock_url_type + ")"
    fig10.update_layout(title=plot_title, xaxis_title='Year', yaxis_title='Rs. Crores')
    
    x = df_balance_sheet.index
    y = df_balance_sheet[df_balance_sheet.columns[0]]  # first trace
    
    fig10.add_traces(go.Bar(x=x, y=y, name='Equity Capital',  marker_color=color_set.get("blue")))

    # create `list` with a `dict` for each column
    buttons = [ {
                 'method': 'update', 'label': col, 
                 'args': [ {'y': [ df_balance_sheet[col] ], 'marker.color': [plot_colors.get(col)]} ]
                } 
                for col in df_balance_sheet.iloc[:, :]
              ]
    # add menus
    updatemenus = [{'buttons': buttons, 'direction': 'down', 'showactive': True,}]
    fig10.update_layout(updatemenus=updatemenus)
    #fig10.show()
    all_figures.append(fig10)
    if(DEBUG_ME): print("Info: Chart-10 Generated")


    ## ==================================================================================================================
    ## Chart-11: Fig11 - [Cash Flows] - `Cash Flow Trends`
    ## ==================================================================================================================

    df_cash_flows = df_cash_flows.rename(columns={'Cash Flows': 'Year'}) # Rename a col
    df_cash_flows = df_cash_flows.set_index('Year').T # Transpose with reset index

    plot_cols = ['Cash from Operating Activity','Cash from Investing Activity','Cash from Financing Activity','Net Cash Flow']

    plot_colors = {
       'Cash from Operating Activity'   : color_set.get("blue"), 
       'Cash from Investing Activity'   : color_set.get("bright_green"),
       'Cash from Financing Activity'   : color_set.get("red"),
       'Net Cash Flow'                  : color_set.get("dark_green"),
    }

    for col in plot_cols:
        if col in df_cash_flows.columns:
            x = df_cash_flows.index
            y = df_cash_flows[col]
            fig11.add_trace(go.Bar(x=x, y=y, name=col, marker_color=plot_colors.get(col))) 

    plot_title = '<b>Cash Flow Trends: </b>' + stock_name + " (" + stock_url_type + ")"

    fig11.update_layout(title=plot_title, xaxis_title='Year', yaxis_title='Rs. Crores')

    updatemenus=[
        dict(
        type="buttons",
        buttons=list([
            dict(label="Group", method="relayout", args=[{"barmode": "group"}]),
            dict(label="Stack", method="relayout", args=[{"barmode": "stack"}])
        ])
     )]
    fig11.update_layout(updatemenus=updatemenus)
    # #fig11.show()
    all_figures.append(fig11)
    if(DEBUG_ME): print("Info: Chart-11 Generated")


    ## ==================================================================================================================
    ## Chart-12: Fig12 - [Ratios] - `ROE or ROCE Trend`
    ## ==================================================================================================================
    plot_title = ''
    
    df_ratios = df_ratios.rename(columns={'Ratios': 'Year'}) # Rename a col
    df_ratios = df_ratios.set_index('Year').T # Transpose with reset index

    if 'ROE %' in df_ratios.columns:
        plot_title = '<b>ROE% Trends: </b>' + stock_name + " (" + stock_url_type + ")"
        x  = df_ratios.index
        y1 = df_ratios['ROE %']
        fig12.add_trace(go.Bar(x=x, y=y1, name='ROE %', marker_color=color_set.get("bright_green")))
        fig12.update_layout(title=plot_title,xaxis_title='Year',yaxis_title='%')
        #fig12.show()
        all_figures.append(fig12)
        if(DEBUG_ME): print("Info: Chart-12 Generated")
    elif 'ROCE %' in df_ratios.columns:
        plot_title = '<b>ROCE% Trends: </b>' + stock_name + " (" + stock_url_type + ")"
        x  = df_ratios.index
        y1 = df_ratios['ROCE %']
        fig12.add_trace(go.Bar(x=x, y=y1, name='ROCE %', marker_color=color_set.get("bright_green")))
        fig12.update_layout(title=plot_title,xaxis_title='Year',yaxis_title='%')
        #fig12.show()
        all_figures.append(fig12)
        if(DEBUG_ME): print("Info: Chart-12 Generated")


    ## ==================================================================================================================
    ## Chart-13: Fig13 - [Ratios] - `Operational Ratios Trend` 
    ## ==================================================================================================================
    plot_title = ''
    plot_cols_all = ['Debtor Days','Inventory Days','Days Payable','Cash Conversion Cycle','Working Capital Days']
    # Not all cols are applicable and present inside df_ratios for specific stock 
    # e.g. For Airlines company, Banks etc, Inventory days is not applicable
    plot_colors = {
       'Debtor Days'           : color_set.get("blue"), 
       'Inventory Days'        : color_set.get("red"),
       'Days Payable'          : color_set.get("orange"),
       'Cash Conversion Cycle' : color_set.get("bright_green"),
       'Working Capital Days'  : color_set.get("purple")
    }

    plot_cols_available = []

    for col in plot_cols_all:
        if col in df_ratios.columns: 
            plot_cols_available.append(col)
    
    if len(plot_cols_available) != 0:
        plot_title = '<b>Operational Ratios : </b>' + stock_name + " (" + stock_url_type + ")"
        # for col in plot_cols_available:
        #     x  = df_ratios.index
        #     y1 = df_ratios[col]
        #     fig12.add_trace(go.Bar(x=x, y=y1, name='ROCE %', marker_color=plot_colors.get(col)))
        fig13.update_layout(title=plot_title,xaxis_title='Year',yaxis_title='%')
        
        x = df_ratios.index
        y = df_ratios[df_ratios.columns[0]]  # first trace
        
        fig13.add_traces(go.Bar(x=x, y=y, name=plot_cols_available[0],  marker_color=plot_colors.get(plot_cols_available[0])))

        # create `list` with a `dict` for each column
        buttons = [ 
            {'method': 'update', 'label': col, 'args': [{'y': [ df_ratios[col] ], 'marker.color': [ plot_colors.get(col) ]}]
            }for col in  plot_cols_available # df_ratios.iloc[:, :]
        ]
        # add menus
        updatemenus = [{'buttons': buttons, 'direction': 'down', 'showactive': True,}]
        fig13.update_layout(updatemenus=updatemenus)
        #fig13.show()
        all_figures.append(fig13)
        if(DEBUG_ME): print("Info: Chart-13 Generated")

    ## ==================================================================================================================
    ## Chart-14: Fig14 - [Shareholding Pattern] - `Participant Shareholding Pattern`
    ## ==================================================================================================================
    df_shareholding_pattern = df_shareholding_pattern.rename(columns={'Shareholding Pattern': 'Quarter'}) # Rename a col
    df_shareholding_pattern = df_shareholding_pattern.set_index('Quarter').T # Transpose with reset index

    plot_cols_all = ['Promoters','FIIs','DIIs','Public','Government']

    plot_colors = {
       'Promoters'  : color_set.get("bright_green"), 
       'FIIs'       : color_set.get("dark_blue"),
       'DIIs'       : color_set.get("blue"),
       'Public'     : color_set.get("red"),
       'Government' : color_set.get("purple")
    }

    for col in plot_cols_all:
        if col in df_shareholding_pattern.columns:
            x = df_shareholding_pattern.index
            y = df_shareholding_pattern[col]
            fig14.add_trace(go.Bar(x=x, y=y, name=col, marker_color=plot_colors.get(col))) 
            
    plot_title = '<b>Shareholding Pattern : </b>' + stock_name + " (" + stock_url_type + ")"
    fig14.update_layout(title=plot_title, xaxis_title='Quarter', yaxis_title='%')
    updatemenus=[
        dict(
            type="buttons",
            buttons=list([
            dict(label="Group", method="relayout", args=[{"barmode": "group"}]),
            dict(label="Stack", method="relayout", args=[{"barmode": "stack"}])
            ]))
    ]
    fig14.update_layout(updatemenus=updatemenus)
    #fig14.show()
    all_figures.append(fig14)
    if(DEBUG_ME): print("Info: Chart-14 Generated")

    ## ==================================================================================================================
    ## Chart-15: Fig15 - [Shareholding Pattern] - `Number of Shareholders`
    ## ==================================================================================================================
    df_shareholding_pattern['Shareholders Change %'] = df_shareholding_pattern['No. of Shareholders'].pct_change()*100
    
    x  = df_shareholding_pattern.index
    y1 = df_shareholding_pattern['No. of Shareholders']
    y2 = df_shareholding_pattern['Shareholders Change %'].apply(lambda x: round(x, 1))
    
    fig15.add_trace( go.Bar( x=x, y=y1, 
                      name='No. of Shareholders', 
                      marker_color=color_set.get("blue"), 
                      text=y2,
                      textposition="outside",
                      texttemplate="%{text} %, [%{y}]"
                     )
                   )

    plot_title = '<b>%Change in Number of Shareholders : </b>' + stock_name + " (" + stock_url_type + ")"
    fig15.update_layout(title=plot_title, xaxis_title='Quarter', yaxis_title=' % ')
    #fig15.show()
    all_figures.append(fig15)
    if(DEBUG_ME): print("Info: Chart-15 Generated")

    ## ==================================================================================================================
    ## Finally, Export all figures to HTML File
    ## ==================================================================================================================

    # Optionally, display all figures
    if(DEBUG_ME): 
        for fig in all_figures: 
            fig.show()

    stock_html_reportpath =  output_folder + "/" + stock_name + "-" + stock_url_type + "-" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S")+".html"

    with open(stock_html_reportpath, 'a') as f:
        for fig in all_figures:
            f.write(fig.to_html(full_html=False, include_plotlyjs='cdn'))
            
    if(DEBUG_ME): print("Info: All charts saved to: ", stock_html_reportpath)
    
    return stock_html_reportpath
## ============================================================================================================

import pandas as pd
import numpy as np
import re
import shutil
import os

input_csv = "Input/stocks-ondemand.csv"

output_report_folder = "Output-OnDemandReport/"
output_archived_folder = "Output-ArchivedReport/"


allfiles = os.listdir(output_report_folder)
# iterate on all files to move them to destination folder
for f in allfiles:
    if f.endswith(".html") or f.endswith(".xlsx"):
        src_path = os.path.join(output_report_folder, f)
        dst_path = os.path.join(output_archived_folder, f)
        shutil.move(src_path, dst_path)
         
df = pd.read_csv(input_csv)


stock_name = ""
stock_url = ""

for index, row in df.iterrows():
    stock_url = row["Screener Stock Url"]
    if m := re.match(r'https:\/\/www\.screener\.in\/company\/([A-Z0-9]+).*', stock_url):
        stock_name = m.group(1).strip()
        xlsx_report_path = ScrapeSingleStock(stock_name,stock_url,output_report_folder)
        print(xlsx_report_path)
        html_report_path = VisualizeSingleStock(xlsx_report_path,output_report_folder)
        print(html_report_path)    
    else: 
        print("Error: Stock name!")

print("Done")

## ============================================================================================================
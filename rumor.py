import json
import requests
import pandas as pd
import os


def set_cookies():
    """
    Set cookies for the request.

    Returns:
    dict: A dictionary of cookies.
    """
    cookies = {
        'wdcid': '555e6438c28f5b58',
        'arialoadData': 'true',
        'ariawapChangeViewPort': 'false',
        'wdlast': '1713760448',
    }
    return cookies


def set_headers():
    """
    Set headers for the request.

    Returns:
    dict: A dictionary of headers.
    """
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Connection': 'keep-alive',
        'Referer': 'https://www.piyao.org.cn/bq/index.htm',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua': '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }
    return headers


def get_resp(cookies, headers):
    """
    Get the JSON response from the specified URL.

    Parameters:
    cookies (dict): Cookies for the request.
    headers (dict): Headers for the request.

    Returns:
    str: The response text.
    """
    response = requests.get(
        'https://www.piyao.org.cn/bq/ds_0615b39e6744461c8af75b06dd4f1c46.json',
        cookies=cookies,
        headers=headers,
    )
    return response.text


def extract_valuable_info(data):
    """
    Extract valuable information from the JSON data.

    Parameters:
    data (dict): The JSON data.

    Returns:
    list: A list of dictionaries containing extracted information.
    """
    extracted_info_list = []

    if 'datasource' in data and len(data['datasource']) > 0:
        for source_item in data['datasource']:
            extracted_info = {}
            extracted_info['title'] = source_item.get('title', 'No title provided')
            extracted_info['summary'] = source_item.get('summary', 'No summary provided')
            extracted_info['publish_time'] = source_item.get('publishTime', 'No publish time provided')

            base_url = "http://piyao.preview.e.news.cn"
            publish_url = source_item.get('publishUrl', '')
            if publish_url.startswith('../'):
                publish_url = publish_url.replace('../', '/')
            full_url = base_url + publish_url
            extracted_info['publish_url'] = full_url

            extracted_info['source'] = source_item.get('sourceText', 'No source provided')
            extracted_info_list.append(extracted_info)

    return extracted_info_list


def create_dataframe(data_info):
    """
    Create a pandas DataFrame from the extracted information.

    Parameters:
    data_info (list): A list of dictionaries containing extracted information.

    Returns:
    pandas.DataFrame: The created DataFrame.
    """
    df = pd.DataFrame(data_info)
    return df


def df_to_excel(df, output_path):
    """
    Save the DataFrame to an Excel file.

    Parameters:
    df (pandas.DataFrame): The DataFrame to save.
    output_path (str): The path to the output Excel file.

    Returns:
    None
    """
    if not os.path.exists(output_path):
        df.to_excel(output_path, index=False)
    else:
        with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            start_row = writer.sheets['Sheet1'].max_row
            header = False if start_row > 1 else True
            df.to_excel(writer, sheet_name='Sheet1', startrow=start_row, index=False, header=header)


def main():
    """
    Main function to execute the data extraction and saving.

    Returns:
    None
    """
    cookies = set_cookies()
    headers = set_headers()

    resp = get_resp(cookies, headers)
    print(resp)
    resp_json = json.loads(resp)

    data = extract_valuable_info(resp_json)

    df = create_dataframe(data)
    df_to_excel(df, '辟谣数据2.xlsx')
    print('Data added to Excel')


if __name__ == '__main__':
    main()

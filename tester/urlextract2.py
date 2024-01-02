import requests

cookies = {
    'ab': 'stage-web-apps',
    '_ga': 'GA1.1.798189167.1703477051',
    '_gcl_au': '1.1.2118757811.1703477052',
    '_fbp': 'fb.1.1703732913826.708081806',
    # '.AspNetCore.Antiforgery.VyLW6ORzMgk': 'CfDJ8KKvyj28jF5Ik7-myJ51XiThun0j6OjyoQeyuImmlypZedyWoBRLRjadyzu2ReoTaLVAFNiJeD-mg6qnDBvXnTzYkHQrkTGXz5wWmIZ58ihXBNDY6jE2dlLGmTCzQDne443QVhFeSwsPGs_HMlpOmto',
    '_ga_YZFGTV3XRZ': 'GS1.1.1704154533.23.1.1704157025.6.0.0',
}

headers = {
    'authority': 'messicks.com',
    'accept': '*/*',
    'accept-language': 'en-US,en;q=0.9,ja-JP;q=0.8,ja;q=0.7,id;q=0.6',
    # 'cookie': 'ab=stage-web-apps; _ga=GA1.1.798189167.1703477051; _gcl_au=1.1.2118757811.1703477052; _fbp=fb.1.1703732913826.708081806; .AspNetCore.Antiforgery.VyLW6ORzMgk=CfDJ8KKvyj28jF5Ik7-myJ51XiThun0j6OjyoQeyuImmlypZedyWoBRLRjadyzu2ReoTaLVAFNiJeD-mg6qnDBvXnTzYkHQrkTGXz5wWmIZ58ihXBNDY6jE2dlLGmTCzQDne443QVhFeSwsPGs_HMlpOmto; _ga_YZFGTV3XRZ=GS1.1.1704154533.23.1.1704157025.6.0.0',
    # 'referer': 'https://messicks.com/vendor/kubota',
    'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'x-requested-with': 'XMLHttpRequest',
}

response = requests.get('https://messicks.com/api/vendor/equipmenttypes/1/0', cookies=cookies, headers=headers)
print(response.text)
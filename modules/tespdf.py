import requests

cookies = {
    '_ga': 'GA1.1.1951540105.1707265469',
    '_gcl_au': '1.1.2035332135.1707265469',
    '_fbp': 'fb.1.1707265473420.2004796248',
    '.AspNetCore.Antiforgery.VyLW6ORzMgk': 'CfDJ8PdKVmDbdbZNuLgh6UCVBNfCn1eNJpgcKPzLdMOEau-Og2Flx0ibgkU_E974XK2USBDhCxCEJ_bwbiOQjow3AkyQjXI5R2sSxAI80PL9LIiITyiEQSRkMraAP7wMyMiLoQ5WfX75MD4SU09yzofWbrc',
    'ab': 'prod-web-apps',
    '_clck': '1eoj49x%7C2%7Cfjp%7C0%7C1521',
    '_ga_YZFGTV3XRZ': 'GS1.1.1709272859.7.1.1709273093.60.0.0',
    '_clsk': 'eg1fi4%7C1709273096358%7C1%7C1%7Cl.clarity.ms%2Fcollect',
}

headers = {
    'authority': 'messicks.com',
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'accept-language': 'en-US,en;q=0.9,id;q=0.8',
    'cache-control': 'no-cache',
    # 'cookie': '_ga=GA1.1.1951540105.1707265469; _gcl_au=1.1.2035332135.1707265469; _fbp=fb.1.1707265473420.2004796248; .AspNetCore.Antiforgery.VyLW6ORzMgk=CfDJ8PdKVmDbdbZNuLgh6UCVBNfCn1eNJpgcKPzLdMOEau-Og2Flx0ibgkU_E974XK2USBDhCxCEJ_bwbiOQjow3AkyQjXI5R2sSxAI80PL9LIiITyiEQSRkMraAP7wMyMiLoQ5WfX75MD4SU09yzofWbrc; ab=prod-web-apps; _clck=1eoj49x%7C2%7Cfjp%7C0%7C1521; _ga_YZFGTV3XRZ=GS1.1.1709272859.7.1.1709273093.60.0.0; _clsk=eg1fi4%7C1709273096358%7C1%7C1%7Cl.clarity.ms%2Fcollect',
    'pragma': 'no-cache',
    'sec-ch-ua': '"Chromium";v="122", "Not(A:Brand";v="24", "Microsoft Edge";v="122"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0',
}

params = {
    'modelid': '86240',
    'diagramid': '451775',
}

response = requests.get('https://messicks.com/diagram/pdf', params=params, cookies=cookies, headers=headers)
data = response.content
with open(f"file1.pdf", 'wb') as s:
    s.write(data)
import requests


t = [
    {'bucket_name': 'new-templates', 'template_name': 'ndf.xlsx', 'output': 'DDE.docx'}
]

res = requests.post('http://localhost:3001/load_templates', json=t)
print(res)
print(res.text)


data = {
    'template_name': 'ndf.xlsx',
    'data': {
        'prix': {
            'deb': 15
        }
    },
    'output_bucket':'new-output',
    'output_name':'test.xlsx'
}

res = requests.post('http://localhost:3001/publipost', json=data)
print(res)
print(res.text)


data = {
    'name':'ndf.xlsx'
}

res = requests.post('http://localhost:3001/documents', json=data)
print(res)
print(res.text)

data = {
    'name':'ndf.xlsx'
}

res = requests.post('http://localhost:3001/documents', json=data)
print(res)
print(res.text)
import requests

url = 'http://localhost:3001'


def configure():
    t = {
        'endpoint':'documents.juniorisep.com',
        'access_key':'minio',
        'passkey':'4wsLZmoo0UMdI5RcdiyNv6St',
    }

    res = requests.post(url+'/configure', json=t)
    print(res)
    print(res.text)

def load_templates():
    print('trying to load')
    t = [
        {'bucket_name': 'new-templates', 'template_name': 'ndf.xlsx'}
    ]

    res = requests.post(url+'/load_templates', json=t)
    print(res)
    print(res.text)


def publipost():
    print('trying to publipost')
    data = {
        'template_name': 'ndf.xlsx',
        'data': {
            'prix': {
                'deb': 15
            }
        },
        'output_bucket': 'new-output',
        'output_name': 'test.xlsx'
    }

    res = requests.post(url+'/publipost', json=data)
    print(res)
    print(res.text)


def get_placeholders():
    print('getting placeholders')

    data = {
        'name': 'ndf.xlsx'
    }

    res = requests.post(url+'/get_placeholders', json=data)
    print(res)
    print(res.text)


if __name__ == '__main__':
    configure()
    load_templates()
    publipost()
    get_placeholders()
    get_placeholders()

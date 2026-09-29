
ایل کامل app.py پروژه OfficialLegalSearch را می‌خواهم در GitHub قرار بدهم.

بله. فایل کامل app.py پروژه OfficialLegalSearch را همین‌جا بفرست/آپلود کن تا برای قرار دادن در GitHub آماده‌اش کنم.

اگر منظورت همان نسخه‌ای است که قبلاً با هم روی آن کار کردیم و الان فایلش را نداری، بگو «همان نسخه قبلی را پیدا کن» تا از فایل‌ها و کارهای قبلی‌ات دنبالش بگردم.


Vale, نسخه کامل

پیداش کردم. این همان نسخه کامل app.py است که می‌خواستید؛ نسخه کامل پروژه با جست‌وجوی آرای قضایی، سامانه قوانین، لغو جست‌وجو و خروجی Word. 


دانلود فایل کامل app.py

برای GitHub فقط بعد از دانلود، نام فایل را دقیقاً به app.py تغییر بدهید و جایگزین app.py فعلی پروژه OfficialLegalSearch کنید.


app_complete_2135.py
Code

LOCKDOWN MODE

No file chosen
Library
/
app_complete_2135.py
app_compl…_2135.py



2072
2073
2074
2075
2076
2077
2078
2079
2080
2081
2082
2083
2084
2085
2086
2087
2088
2089
2090
2091
2092
2093
2094
2095
2096
2097
2098
2099
2100
2101
2102
2103
2104
2105
2106
2107
2108
2109
2110
2111
2112
2113
2114
2115
2116
2117
2118
2119
2120
2121
2122
2123
2124
2125
2126
2127
2128
2129
2130
2131
2132
2133
2134
2135
            job['completed_pages'],

        total_pages=
            job['total_pages'],

        progress=
            job['progress'],

        message=
            job['message']
    )


@app.post('/api/cancel/<jid>')
def cancel(jid):

    job = JOBS.get(
        jid
    )

    if not job:

        return jsonify(
            error='یافت نشد'
        ), 404

    job['cancel'] = True

    return jsonify(
        ok=True
    )


@app.get('/api/download/<jid>')
def download(jid):

    if jid not in JOBS:

        return (
            'Not found',
            404
        )

    return send_file(
        make_doc(jid),

        as_attachment=True,

        download_name=(
            'national-judicial-decisions.docx'
        )
    )


if __name__ == '__main__':
    app.run(
        host='0.0.0.0',
        port=int(
            os.getenv(
                'PORT',
                5000
            )
        )
    )

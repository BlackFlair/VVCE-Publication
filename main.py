from flask import Flask, url_for, render_template, request, session, redirect, g, Response
import sqlite3
import os, io, xlwt

app = Flask(__name__)
app.secret_key = os.urandom(24)

currentDirectory = os.path.dirname(os.path.abspath(__file__))

db = "\Publication.db"

global s_result

@app.route('/download')
def download():
    global s_result

    if g.user:
        output = io.BytesIO()

        workbook = xlwt.Workbook()

        sh = workbook.add_sheet('Search Result')

        sh.write(0, 0, 'Paper Id')
        sh.write(0, 1, 'Paper Title')
        sh.write(0, 2, 'Citation Number')
        sh.write(0, 3, 'Journal Name')
        sh.write(0, 4, 'Publication Link')

        i = 1
        for row in s_result:
            for j in range (len(row)):
                sh.write(i, j, row[j])
            i+=1

        workbook.save(output)
        output.seek(0)

        return Response(output, mimetype="application/ms-excel", headers={"Content-Disposition":"attachment;filename=Publications_Search_Result.xls"})


@app.route('/', methods=['POST', 'GET'])
def index():
    session.pop('user', None)

    if request.method == 'POST':
        session.pop('user', None)

        connection = sqlite3.connect(currentDirectory + db)
        cursor = connection.cursor()

        query = '''SELECT Password FROM Login WHERE Name = "{n}"'''.format(n=request.form['userName'])

        try:
            pwd = cursor.execute(query)
            pwd = pwd.fetchone()[0]
            connection.close()

        except:
            print("User name does not exist.")  # Display invalid login message

        if request.form['password'] == pwd:
            session['user'] = request.form['userName']  # userName : name given in html component for user name field
            return redirect(url_for('home'))
        else:
            print("Invalid Password.")

    return render_template('signin.html')


@app.route('/home', methods=['POST', 'GET'])
def home():
    connection = sqlite3.connect(currentDirectory + db)
    cursor = connection.cursor()

    query1 = '''SELECT PaperID, PaperTitle, CitationNumber, PublishedIn, PublicationLink, DateOfPublication FROM Publications WHERE UserName="{name}"'''.format(name=g.user)

    result = cursor.execute(query1).fetchall()
    global s_result
    s_result = result

    if g.user:
        print(result)

        row = len(result)

        search_message = ""

        if request.form.get('search_button'):
            print("Button pressed.")
            search_paperTitle = request.form['search_paperTitle']
            search_facultyAuthor = request.form['search_facultyAuthor']
            search_journalName = request.form['search_journalName']
            search_journalType = request.form['search_journalType']
            search_citationNumber = request.form['search_citationNumber']

            condition = ""

            x = [search_paperTitle, search_facultyAuthor, search_journalName, search_journalType, search_citationNumber, '']
            y = ['PaperTitle=', ' FacultyAuthor=', ' PublishedIn=', ' Level=', ' CitationNumber=', None]
            z = ["'{paperTitle}'", "'{facultyAuthor}'", "'{journalName}'", "'{journalType}'", "'{citationNumber}'", None]

            for i in range(len(x)-1):
                if x[i] != "":
                    condition += y[i]+z[i]
                    print(condition)
                if x[i + 1] != "" :
                    condition += "and"
                    print(x[i+1])


            query2 = '''SELECT PaperID, PaperTitle, CitationNumber, PublishedIn, PublicationLink, DateOfPublication FROM Publications WHERE '''+condition.format(paperTitle=search_paperTitle,facultyAuthor=search_facultyAuthor,journalName=search_journalName,journalType=search_journalType,citationNumber=search_citationNumber)

            print(query2)

            search_message = "Showing results for: " + condition

            try:
                result = cursor.execute(query2).fetchall()
                s_result = result
            except sqlite3.Error as e:
                search_message = e
            print(result)

            row = len(result)

        return render_template('home.html', user=g.user, result=result, row=row, search_message=search_message)

    return redirect(url_for('index'))


@app.route('/signup', methods=['POST', 'GET'])
def signup():
    if request.method == 'POST':
        session.pop('user', None)

        connection = sqlite3.connect(currentDirectory + db)
        cursor = connection.cursor()

        query1 = '''INSERT INTO Login VALUES(?, ?)'''

        try:
            userName = request.form['userName']
            password = request.form['password']
            conPass = request.form['conPass']

            if password == conPass:
                print("Password matched")
                cursor.execute(query1, (userName, password))
                print("Insert successful")
                connection.commit()
                print("Commit successful")
                connection.close()
                print("Close successful")

                return render_template('signin.html')

            else:
                print("Password entered does not match.")

        except:
            print("User Name already exists.")

    return render_template('signup.html')


@app.route('/addPublication', methods=['POST', 'GET'])
def addPublication():
    if g.user:
        connection = sqlite3.connect(currentDirectory + db)
        cursor = connection.cursor()

        query1 = '''SELECT Salutation, Name FROM Faculty'''
        query1 = cursor.execute(query1)
        query1 = query1.fetchall()

        length = len(query1)

        fAuthor = []

        print(query1)
        print(g.user)

        for i in range(length):
            name = query1[i][0] + " " + query1[i][1]
            fAuthor.append(name)
            print(fAuthor)

        if request.method == 'POST':

            query2 = '''INSERT INTO Publications 
                    (PaperTitle, At, FacultyAuthor, StudentAuthor, Abstract,
                    PublishedIn, Level, DateOfPublication, PublicationIndex, ISSN_ISBN,
                    DOI, PublicationLink, UploadLink, CertificateLink, ImpactFactorOfPublication, 
                    Cited, CitationNumber, HIndex, FinancialAssistance, Amount, UserName)
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''

            try:
                paperTitle = request.form['paperTitle']
                checkbox = ", ".join(request.form.getlist('checkbox'))
                facultyAuthor = request.form.get('facultyAuthor')
                studentAuthor = request.form['studentNames']
                paragraphText = request.form['paragraphText']
                publishedIn = request.form['publishedIn']
                journal = request.form.get('journal')
                date = request.form['date']
                publicationIndex = request.form['publicationIndex']
                ISSN = request.form.get('ISSN_ISBN')
                publicationLink = request.form['publicationLink']
                uploadLink = request.form['uploadLink']
                certificationLink = request.form['certificationLink']
                impactFactor = request.form['impactFactor']
                doi = request.form['doi']
                cited = request.form.get('cited')
                citedNumber = request.form['citedNumber']
                hIndex = request.form['hIndex']
                assistance = request.form.get('assistance')
                amount = request.form['amount']
                print("Variable assignment successful")

                cursor.execute(query2, (paperTitle, checkbox, facultyAuthor, studentAuthor, paragraphText,
                                        publishedIn, journal, date, publicationIndex, ISSN, doi,
                                        publicationLink, uploadLink, certificationLink, impactFactor,
                                        cited, citedNumber, hIndex, assistance, amount, g.user))
                print("Insert successful")

                connection.commit()
                print("Commit successful")

                connection.close()
                print("Close successful")

                return redirect(url_for('home'))

            except sqlite3.Error as e:
                print("Data entered error")
                print(e)

        return render_template('addPublication.html', length=length, fAuthor=fAuthor)

    return render_template('signin.html')


@app.before_request
def beforeRequest():
    g.user = None

    if 'user' in session:
        g.user = session['user']


@app.route('/dropSession')
def dropSession():
    session.pop('user', None)

    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)

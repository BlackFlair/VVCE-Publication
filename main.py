from flask import Flask, url_for, render_template, request, session, redirect, g, Response
import sqlite3
import os, io, xlwt

currentDirectory = os.path.dirname(os.path.abspath(__file__))

db = "\Publication.db"

app = Flask(__name__)
app.secret_key = os.urandom(24)

global s_result

@app.route('/download')
def download():
    global s_result
    _result = s_result.split()

    connection = sqlite3.connect(currentDirectory + db)
    cursor = connection.cursor()

    if g.user:
        output = io.BytesIO()

        workbook = xlwt.Workbook()

        sh = workbook.add_sheet('Search Result')

        headding = ['Paper ID', 'Paper Title', 'At', 'Faculty Author', 'Student Author', 'Abstract', 'Published In', 'Level', 'Date of Publication', 'Index', 'ISSN/ISBN', 'DOI', 'Publication Link', 'Upload Link', 'Certification Link', 'Impact Factor', 'Cited', 'Citation Number', 'H-Index', 'Financial Assistance', 'Amount', 'User Name']

        for i in range(len(headding)):
            sh.write(0, i, headding[i])

        for i in range(len(_result)):
            query1 = '''SELECT * FROM Publications WHERE PaperID = {id}'''.format(id=int(_result[i]))
            query1 = cursor.execute(query1).fetchall()
            j=0
            for x in range(len(query1[0])):
                sh.write(i+1, j, query1[0][x])
                j+=1

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

        except sqlite3.Error as e:
            print(e)  # Display invalid login message

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

    query1 = '''SELECT PaperID, PaperTitle, _Index, PublishedIn, PublicationLink, DateOfPublication FROM Publications WHERE UserName="{name}"'''.format(name=g.user)

    query3 = '''SELECT Role FROM Login WHERE Name="{name}"'''.format(name=g.user)

    try:
        role = cursor.execute(query3).fetchall()
        result = cursor.execute(query1).fetchall()
    except sqlite3.Error as e:
        print(e)


    _result = ''
    for i in range(len(result)):
        _result = _result + " " + str(result[i][0])

    global s_result
    s_result = _result

    if g.user:
        row = len(result)

        search_message = ""

        if 'search' in request.form:
            search_paperTitle = request.form['search_paperTitle']
            search_facultyAuthor = request.form['search_facultyAuthor']
            search_journalName = request.form['search_journalName']
            search_journalType = request.form['search_journalType']

            # search_from = request.form['from']
            # search_to = request.form['to']
            #
            # dateRange = ''
            #
            # if search_from == '':
            #     pass
            # else:
            #     pass
            #
            # if search_to == '':
            #     pass
            # else:
            #     pass

            condition = ""

            x = [search_paperTitle, search_facultyAuthor, search_journalName, search_journalType, '']
            y = ['PaperTitle=', ' FacultyAuthor=', ' PublishedIn=', ' Level=',  None]
            z = ["'{paperTitle}'", "'{facultyAuthor}'", "'{journalName}'", "'{journalType}'",  None]

            for i in range(len(x)-1):
                if x[i] != "":
                    condition += y[i]+z[i]
                if x[i + 1] != "" :
                    condition += "and"

            query2 = '''SELECT PaperID, PaperTitle, _Index, PublishedIn, PublicationLink, DateOfPublication FROM Publications WHERE '''+condition.format(paperTitle=search_paperTitle,facultyAuthor=search_facultyAuthor,journalName=search_journalName,journalType=search_journalType)

            try:
                result = cursor.execute(query2).fetchall()
                _result = ''
                for i in range(len(result)):
                    _result = _result + " " + str(result[i][0])
                s_result = _result
            except sqlite3.Error as e:
                search_message = e

            row = len(result)

        elif 'edit' in request.form:
            id = request.form['edit']
            id = id[5:]

            connection = sqlite3.connect(currentDirectory + db)
            cursor = connection.cursor()

            query1 = '''SELECT PaperID, PaperTitle, At, FacultyAuthor, StudentAuthor, Abstract,
                                        PublishedIn, Level, DateOfPublication, _Index, ISSN_ISBN,
                                        DOI, PublicationLink, UploadLink, CertificateLink, ImpactFactorOfPublication, 
                                        Cited, CitationNumber, HIndex, FinancialAssistance, Amount, UserName 
                                        FROM Publications WHERE PaperID={id}'''.format(id=id)

            query1 = cursor.execute(query1)
            query1 = query1.fetchall()

            query2 = '''SELECT Salutation, Name FROM Faculty'''
            query2 = cursor.execute(query2)
            query2 = query2.fetchall()

            length = len(query2)

            fAuthor = []

            for i in range(length):
                name = str(query2[i][0]) + " " + str(query2[i][1])
                fAuthor.append(name)

            return render_template('editPublication.html', length=length, fAuthor=fAuthor, query1=query1)

        elif 'delete' in request.form:
            query4 = '''DELETE FROM Publications WHERE PaperID=?'''

            id = request.form['delete']
            id = id[7:]

            cursor.execute(query4, (id,))
            connection.commit()

            return redirect(url_for('home'))

        return render_template('home.html', user=g.user, result=result, row=row, search_message=search_message, role=role)

    return redirect(url_for('index'))


@app.route('/editPublication', methods=['POST', 'GET'])
def editPublication():
    if g.user:
        connection = sqlite3.connect(currentDirectory + db)
        cursor = connection.cursor()

        if request.method == 'POST':
            id = request.form['id']
            query3 = '''UPDATE Publications SET
                                PaperTitle=?, At=?, FacultyAuthor=?, StudentAuthor=?, Abstract=?,
                                PublishedIn=?, Level=?, DateOfPublication=?, _Index=?, ISSN_ISBN=?,
                                DOI=?, PublicationLink=?, UploadLink=?, CertificateLink=?, ImpactFactorOfPublication=?, 
                                Cited=?, CitationNumber=?, HIndex=?, FinancialAssistance=?, Amount=?
                                WHERE PaperID=?'''

            try:
                paperTitle = request.form['paperTitle']
                checkbox = ", ".join(request.form.getlist('checkbox'))
                facultyAuthor = request.form.get('facultyAuthor')
                studentAuthor = request.form['studentNames']
                paragraphText = request.form['paragraphText']
                publishedIn = request.form['publishedIn']
                journal = request.form.get('journal')
                date = request.form['date']
                index = request.form.get('index')
                if (index == "Other"):
                    index = request.form['index']
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

                cursor.execute(query3, (paperTitle, checkbox, facultyAuthor, studentAuthor, paragraphText,
                                        publishedIn, journal, date, index, ISSN, doi,
                                        publicationLink, uploadLink, certificationLink, impactFactor,
                                        cited, citedNumber, hIndex, assistance, amount, id))

                connection.commit()
                connection.close()

                return redirect(url_for('home'))

            except sqlite3.Error as e:
                print(e)

    return render_template('signin.html')

@app.route('/signup', methods=['POST', 'GET'])
def signup():
    if request.method == 'POST':
        session.pop('user', None)

        connection = sqlite3.connect(currentDirectory + db)
        cursor = connection.cursor()

        query1 = '''INSERT INTO Login VALUES(?, ?)'''
        query2 = '''INSERT INTO Faculty (ID) VALUES (?)'''


        try:
            userName = request.form['userName']
            password = request.form['password']
            conPass = request.form['conPass']

            if password == conPass:
                cursor.execute(query1, (userName, password))
                connection.commit()
                cursor.execute(query2, (userName,))
                connection.commit()
                cursor.close()

                return render_template('signin.html')

            else:
                print("Password entered does not match.")

        except sqlite3.Error as e:
            print(e)

        connection.close()

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

        for i in range(length):
            name = str(query1[i][0]) + " " + str(query1[i][1])
            fAuthor.append(name)

        if request.method == 'POST':
            query2 = '''INSERT INTO Publications 
                    (PaperTitle, At, FacultyAuthor, StudentAuthor, Abstract,
                    PublishedIn, Level, DateOfPublication, _Index, ISSN_ISBN,
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
                index = request.form.get('index')
                if (index == "Other"):
                    index = request.form['index']
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

                cursor.execute(query2, (paperTitle, checkbox, facultyAuthor, studentAuthor, paragraphText,
                                        publishedIn, journal, date, index, ISSN, doi,
                                        publicationLink, uploadLink, certificationLink, impactFactor,
                                        cited, citedNumber, hIndex, assistance, amount, g.user))

                connection.commit()
                connection.close()

                return redirect(url_for('home'))

            except sqlite3.Error as e:
                print(e)

        return render_template('addPublication.html', length=length, fAuthor=fAuthor)

    return render_template('signin.html')


@app.route('/facultyDetails', methods=['POST', 'GET'])
def facultyDetails():

    if g.user:
        connection = sqlite3.connect(currentDirectory + db)
        cursor = connection.cursor()

        query1 = '''SELECT * FROM JobRole'''

        query1 = cursor.execute(query1)
        query1 = query1.fetchall()

        jobRole_length = len(query1)

        jobRole = []

        for i in range(jobRole_length):
            jobRole.append(query1[i][1])

        query2 = '''SELECT * FROM Faculty WHERE ID="{name}"'''.format(name=g.user)

        query2 = cursor.execute(query2)
        query2 = query2.fetchall()

        if request.method == "POST":
            try:
                id = request.form['id']
                salutation = request.form.get('salutation')
                name = request.form['name']
                phone = request.form['phone']
                email = request.form['email']
                pan = request.form['pan']
                panImg = request.files['panImg']
                aadhar = request.form['aadhar']
                aadharImg = request.files['aadharImg']
                accNum = request.form['bank']
                ifsc = request.form['ifsc']
                dojCurrent = request.form['doj_current']
                department = request.form['department']
                designation = request.form['designation']
                promoted = request.form.get('promoted')
                promotionOrder = request.files['promotionOrder']
                appointmentOrder = request.files['appointmentOrder']
                phd = request.form.get('phd')
                phdDate = request.form['phdDate']
                contract = request.form.get('contract')
                adjunct = request.form.get('adjunct')
                degree = request.form['degree']
                college = request.form['college']
                university = request.form['university']
                yob = request.form['yob']
                yoc = request.form['yoc']
                certificate = request.files['certificate']
                organization = request.form['organization']
                orgDesignation = request.form['org_designation']
                orgDoj = request.form['org_doj']
                orgDor = request.form['org_dor']
                orgReleavingLetter = request.files['org_relievingLetter']
                orgStatus = request.form.get('org_status')

                panImg = panImg.read()
                aadharImg = aadharImg.read()
                certificate = certificate.read()
                promotionOrder = promotionOrder.read()
                appointmentOrder = appointmentOrder.read()
                orgReleavingLetter = orgReleavingLetter.read()

                educationHistory = {'Degree':degree, 'College':college, 'University':university, 'Joining':yob, 'Completion':yoc, 'Certificate': certificate}
                workHistory = {'Organization':organization, 'Designation':orgDesignation, 'DOJ':orgDoj, 'DOR':orgDor, 'ReleavingLetter':orgReleavingLetter, 'Status':orgStatus}

                query3 = '''INSERT INTO Faculty (ID, Salutation, Name, Phone, Email, PANNumber, PANImage, 
                            AadharNumber, AadharImage, AccountNumber, IFSC, DOJ, Designation, Department, 
                            Promoted, PromotionOrder, RegisteredPhD, PhDRegDate, AppointmentOrder, 
                            Contract, AdjunctFaculty, Education, WorkHistory) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? )'''

                query4 = '''DELETE FROM Faculty WHERE ID="{id}"'''.format(id=id)

                cursor.execute(query4)

                cursor.execute(query3, (id, salutation, name, phone, email, pan, panImg, aadhar, aadharImg, accNum,
                                                    ifsc, dojCurrent, designation, department, promoted, promotionOrder,
                                                    phd, phdDate, appointmentOrder, contract, adjunct, str(educationHistory), str(workHistory)))

                connection.commit()

            except sqlite3.Error as e:
                print(e)

        return render_template('facultyDetails.html', jobRole=jobRole, details=query2)
    
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

import win32com.client, os, xlrd


Wordpath = "testd_old.docx"
xlspath = "testd.xlsx"
absWordPath = os.path.abspath(Wordpath)
absExcelPath = os.path.abspath(xlspath)
#print(absPath)

wdFindContinue = 1
wdReplaceAll = 2

#app = win32com.client.Dispatch('Word.Application')
app = win32com.client.DispatchEx("Word.Application")
app.Visible = 0
app.DisplayAlerts = 0
app.Documents.Open(absWordPath)


def search_replace_all(word_file, find_str, replace_str):
    ''' replace all occurrences of `find_str` w/ `replace_str` in `word_file` '''
    wdFindContinue = 1
    wdReplaceAll = 2

    # Dispatch() attempts to do a GetObject() before creating a new one.
    # DispatchEx() just creates a new one.
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = 0
    app.DisplayAlerts = 0
    app.Documents.Open(word_file)

    # expression.Execute(FindText, MatchCase, MatchWholeWord,
    #   MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward,
    #   Wrap, Format, ReplaceWith, Replace)
    app.Selection.Find.Execute(find_str, False, False, False, False, False,
        True, wdFindContinue, False, replace_str, wdReplaceAll)
    app.ActiveDocument.Close(SaveChanges=True)
    app.Quit()


xlsdata = xlrd.open_workbook(absExcelPath)
table = xlsdata.sheets()[0]
nrows = table.nrows
ncols = table.ncols
for j in range(1, 4):
    k = 1
    while k < ncols:
        for i in range(nrows):
            flag = False
            # str = table.cell(i,1)
            strval = table.cell(i, k).value
            # print(str)
            if (strval != ''):
                print(i)
                if (isinstance(table.cell(i, k + 1).value, float)):
                    if (table.cell(i, k + 1).value < 1):
                        restr = str(table.cell(i, k + 1).value)
                        flag = True
                if (flag == False):
                    restr = str(int(table.cell(i, k + 1).value) * j)
                restr = strval + restr + "克"
                strval = strval + str(j) + "袋"
                print(strval, restr)
                # search_replace_all(absWordPath, strval, restr)
                app.Selection.Find.Execute(strval, False, False, False, False, False,
                                           True, wdFindContinue, False, restr, wdReplaceAll)
            else:
                print(i)
        k += 2


app.ActiveDocument.Close(SaveChanges=True)
app.Quit()
'''
doc = app.Documents.Open(absPath)

app.Visible = True
app.ScreenUpdating = True

#doc.Content.Find.Execute(FindText=u'123', ReplaceWith=u'aaa', Replace=2, Wrap=1)


find=app.Selection.Find
find.Wrap = 1
find.ClearFormatting()
find.Text=u'123'
find.Replacement.ClearFormatting()
find.Replacement.Text = u'12re'

find.Execute(Replace=2, Forward=True)

#doc.Content.Find.Execute(FindText=u'abcd', ReplaceWith=u'1234', Replace=2)
'''
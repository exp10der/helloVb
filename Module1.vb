Imports System.Text
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Runtime.CompilerServices
Imports HtmlAgilityPack

Module Module1
    Const base As String = "http://www.cds.spb.ru"
    Sub Main()
        ' Результирующий словарь
        Dim map As New Dictionary(Of String, String)



        Dim realEstates = Parse("http://www.cds.spb.ru/novostroiki-peterburga/")

        Dim cantale = realEstates.Find(Function(n) n.Title.Contains("Кантеле"))
        If cantale IsNot Nothing Then
            realEstates.Remove(realEstates.Find(Function(n) n.Title.Contains("Кантеле")))
            realEstates.AddRange(Parse(cantale.Uri))
        End If



        ' Проходим по списку на странице новостройки и заполняем корпуса
        For Each estate As RealEstate In realEstates
            estate.HousingEstates = New List(Of HousingEstate)()
            Dim site = GetHtml(estate.Uri)

            Dim htmldoc = New HtmlDocument()
            htmldoc.LoadHtml(site)
            Dim housing = htmldoc.DocumentNode.SelectNodes(".//*[@id='outer']/div[2]/div/table/tr/td[2]/div/div/div[2]/a")
            If housing IsNot Nothing Then
                For Each htmlNode In housing
                    Try
                        Dim urlsite = GetFixUrl(htmlNode.GetAttributeValue("href", ""))
                        'Получили адресс теперь необходимо проверить на якори если сылка с якорем то пропускаем итерацию
                        If urlsite.Contains("#") Then
                            Continue For
                        End If
                        map.Add(String.Format("{0} {1} {2}", estate.Title, estate.ShortAddress, htmlNode.SelectSingleNode("div/span").InnerText.Trim()), urlsite)
                    Catch ex As Exception
                        Continue For
                    End Try

                Next
            Else
                map.Add(String.Format("{0} {1}", estate.Title, estate.ShortAddress), estate.Uri)
            End If
        Next



    End Sub

    ''' <summary>
    ''' Получаем исходный код страницы
    ''' </summary>
    ''' <param name="uri">Адресс страницы</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetHtml(uri As String) As String
        Dim html As String

        Dim http As HttpWebRequest = WebRequest.Create(uri)
        http.KeepAlive = False 'True
        http.Proxy = Nothing
        http.Referer = base

        Dim httpr As HttpWebResponse = http.GetResponse()
        '20866 Кодировка koi8
        Using stream = New StreamReader(httpr.GetResponseStream(), Encoding.GetEncoding(20866))
            html = stream.ReadToEnd()
        End Using
        Return html
    End Function

    Private Function Parse(uri As String) As List(Of RealEstate)
        Dim code = GetHtml(uri)

        Dim doc = New HtmlDocument
        doc.LoadHtml(code)
        ' ".//*[@id='outer']/div[2]/div/table/tbody/tr/td[2]/div/div/div[2]/a"
        Dim node = doc.DocumentNode.SelectNodes(".//*[@id='outer']/div[2]/div/table/tr/td[2]/div[2]/a")
        ' Если нода null то это говорит о том что изменилась верстка страницы
        If node Is Nothing Then
            Throw New Exception("Ошибка при парсинге страницы")
        End If

        Dim list = New List(Of RealEstate)
        For Each htmlNode In node
            Try
                Dim titleTmp = htmlNode.SelectSingleNode("header/h2").InnerText.Trim().FixString()

                ' Этим условием убираем последнии квартиры в сданных домах
                If String.IsNullOrWhiteSpace(titleTmp) Then
                    Continue For
                End If

                Dim shortAddressTmp = htmlNode.SelectSingleNode("header/p/em")
                Dim uriTmp = htmlNode.GetAttributeValue("href", "")

                list.Add(New RealEstate() With {
                                .Title = titleTmp,
                                .ShortAddress = If(shortAddressTmp IsNot Nothing, shortAddressTmp.InnerText.Trim().FixString(), ""),
                                .Uri = GetFixUrl(uriTmp)
                     })
            Catch ex As Exception
                Continue For
            End Try
        Next

        Return list

    End Function
    'Исправляем адресс к странице до абсолютного
    Private Function GetFixUrl(ByVal currentUri As String) As String

        Return If(New Uri(currentUri, UriKind.RelativeOrAbsolute).IsAbsoluteUri, currentUri, "http://www.cds.spb.ru" & Convert.ToString(currentUri))
    End Function
End Module

Class RealEstate
    ''' <summary>
    ''' Короткое название
    ''' </summary>
    Public Property Title As String

    ''' <summary>
    ''' Адрес
    ''' </summary>
    Public Property ShortAddress As String

    ''' <summary>
    ''' Сылка на Жилой комплекс
    ''' </summary>
    Public Property Uri As String

    ''' <summary>
    ''' Тут идет вторая вложенность страницы где название корпусов
    ''' некотрые страницы могут иметь сразу сылку на табличку для квартир
    ''' </summary>
    Public Property HousingEstates As List(Of HousingEstate)
End Class
Class HousingEstate
    ' Название корпуса
    Public Property NameHousing As String
    ' Сылка на таблицу
    Public Property Uri As String
End Class


Module StringExtensions

    'Выпиливаем спец символы
    <Extension()>
    Public Function FixString(ByVal str As String)
        Return Regex.Replace(str, "[a-zA-z&;<>]*", "")
    End Function
End Module
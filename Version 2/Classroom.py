from bs4 import BeautifulSoup
import requests

class Classroom():
    def __init__(self, id):
        self.id = id
        self.html = None
        self.room = None
        self.instructor = None

    def grabClassPage(self):
        url = "https://www.inova.org/creg/classdetails.aspx?sid=1&ClassID=%d&sslRedirect=true"
        header = {
            'User-Agent': 'Mozilla/5.0 (X11; OpenBSD i386) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1985.125 Safari/537.36}'}
        htmlContent = requests.get(url % self.id, headers=header, timeout=60, allow_redirects=True)
        if htmlContent.status_code != 200:
            return False
        soup = BeautifulSoup(htmlContent.text, 'html.parser')
        self.html = soup
        return True

    def grabRoom(self):
        roomText = self.html.find(text="Room: ")
        if roomText == None:
            return
        tag = roomText.parent
        room = tag.next_sibling
        self.room = room

    def grabInstructor(self):
        instructText = self.html.find(text="Instructor: ")
        if instructText == None:
            return
        tag = instructText.parent
        instructor = tag.next_sibling
        self.instructor = instructor

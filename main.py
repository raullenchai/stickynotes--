#!/usr/bin/env python
import os
import urllib

from google.appengine.api import users
from google.appengine.ext import blobstore
from google.appengine.ext import db
from google.appengine.ext import webapp
from google.appengine.ext.webapp import blobstore_handlers
from google.appengine.ext.webapp import template
from google.appengine.ext.webapp.util import run_wsgi_app

import OleFileIO_PL
from StringIO import StringIO
import Rtf2Txt


class UserNote(db.Model):
  user = db.UserProperty()
  machine = db.StringProperty(default='NULL')
  blob_key = blobstore.BlobReferenceProperty()
  time = db.DateTimeProperty(auto_now_add=True)
  reserved1 = db.StringProperty(default='NULL')
  reserved2 = db.StringProperty(default='NULL')
  reserved3 = db.IntegerProperty(default=0)
  reserved4 = db.IntegerProperty(default=0)
  
  
class MainHandler(webapp.RequestHandler):
    def get(self):
        upload_url = blobstore.create_upload_url('/upload')
        self.response.out.write('<html><body>')
        self.response.out.write('<form action="%s" method="POST" enctype="multipart/form-data">' % upload_url)
        self.response.out.write("""Upload File: <input type="file" name="file"><br> <input type="submit"
            name="submit" value="Submit"> </form></body></html>""")

class NoteUploadHandler(blobstore_handlers.BlobstoreUploadHandler):
    def post(self):
        try:
            upload = self.get_uploads()[0]
            tmp_key = upload.key()
            user_note = UserNote(user=users.get_current_user(), blob_key=upload.key())
            db.put(user_note)	#store it in GAE database

            blob_reader = blobstore.BlobReader(tmp_key, buffer_size=1048576)
            value = blob_reader.read()
            file = StringIO(value)
            oleobj = OleFileIO_PL.OleFileIO(file)
            assert oleobj.openstream('Version').read() == '02000000'.decode('hex')  # make sure StickyNotes.snt comes from Win 7

            for entry in sorted(oleobj.listdir()):
                if len(entry) == 2 and entry[1] == '0':
                    rtf = oleobj.openstream(entry).read()
                    rtf = rtf[:rtf.index('\0')]
                    rtf = rtf.replace(r'\ansicpg936', '')	# TODO: multiple language support
                    txt = Rtf2Txt.getTxt(rtf)
                    print txt
            print tmp_key
        except:
            self.redirect('/upload_failure.html')

class QueryHandler(webapp.RequestHandler):
  def get(self):
    usernotes = UserNote.all()
    #usernotes.filter("user =", 'raullen')
    usernotes.order("time")
    for usernote in usernotes:        
        print usernote.blob_key
        blob_reader = blobstore.BlobReader(usernote.blob_key, buffer_size=1048576)
        value = blob_reader.read()
        file = StringIO(value)
        ole = OleFileIO_PL.OleFileIO(file)
        assert ole.openstream('Version').read() == '02000000'.decode('hex')  # make sure StickyNotes.snt comes from Win 7

        for entry in sorted(ole.listdir()):
            if len(entry) == 2 and entry[1] == '0':
                rtf = ole.openstream(entry).read()
                rtf = rtf[:rtf.index('\0')]
                rtf = rtf.replace(r'\ansicpg936', '')  # no one uses Chinese /_\
                txt = Rtf2Txt.getTxt(rtf)
                print txt
    #self.redirect('/serve/%s' % blob_info.key())

class ServeHandler(blobstore_handlers.BlobstoreDownloadHandler):
  def get(self, resource):
    resource = str(urllib.unquote(resource))
    blob_info = blobstore.BlobInfo.get(resource)
    blob_reader = blobstore.BlobReader('KSyh5UXl7TSETnE-vaSGYw==', buffer_size=1048576)
    value = blob_reader.read()
    file = StringIO(value)
    ole = OleFileIO_PL.OleFileIO(file)
    assert ole.openstream('Version').read() == '02000000'.decode('hex')  # make sure StickyNotes.snt comes from Win 7

    for entry in sorted(ole.listdir()):
        if len(entry) == 2 and entry[1] == '0':
            rtf = ole.openstream(entry).read()
            rtf = rtf[:rtf.index('\0')]
            #rtf = rtf.replace(r'\ansicpg936', '')  # no one uses Chinese /_\
            txt = Rtf2Txt.getTxt(rtf)
            print txt


def main():
  application = webapp.WSGIApplication(
    [('/', MainHandler),
     ('/upload', NoteUploadHandler),
     ('/serve/([^/]+)?', ServeHandler),
	 ('/query', QueryHandler),
    ], debug=True)
  run_wsgi_app(application)

if __name__ == '__main__':
  main()
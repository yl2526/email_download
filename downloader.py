# -*- coding: utf-8 -*-
"""
Created on Thu Dec 03 13:33:49 2015

@author: yliu

This scrit is to create a email downloader object

The object will create a imap connection. Then you can pass the list of attachmet 
or certaim type of string to download.

"""

import os
import getpass
import imaplib
import email
import re

class downloader:
    '''
    class for email attachment downloading
    '''
    def __init__(self, userName = None, passWord = None, server = 'imap.gmail.com', folder = 'Inbox', search = 'All', target = 'All'):
        '''
        initialize downloader with following:
            server:     default to imap server address 
            username:   default to None, no '@somethong.com' part
            passowrd:   default to None
            login_now:  default to True, is or not login immediately
            search:     search rules default to 'ALL' if '' pass
            target:     target attachment list if default to None, all attachment will be downloaded
            folder:     folder of email to look
        
        returns a downloader with follosing attributes:
            _server:            imap server address to connect
            _username:          username of the individual email
            _imap:              imap connection to the server
            _search:            current search rules, default to 'All'
            _target:            list of dicts, each have subject, attachment keys to store names, 
                                default string 'All' will download all attachments 
            _folder:            name of folder to select
        '''
        self._server = server
        self._imap = imaplib.IMAP4_SSL(server)
        if userName == None:
            self._username = raw_input('Enter your username:')
        else:
            self._username = userName        
        if passWord == None:
            passWord = getpass.getpass('Enter your password: ')
        login_result = self._imap.login(self._username, passWord)
        assert login_result[0] == 'OK', 'unable to login for user {0}, at server {1}'.format(self._userName, self._server)
        self._search = search
        self._target = target
        self._folder = folder
        result, _ = self._imap.select(folder)
        assert (result == 'OK'), 'unable to select the {0} folder'.format(folder)
    
    def __repr__(self):
        '''
        the general representaiton
        '''
        return "User: {0} Server: {1} Folder: {2}".format(self._username, self._server, self._folder)

    def __str__(self):
        '''
        generate string description for downloader.
        '''
        description = 'This downloader connects to server: {0}\n'.format(self._server)
        description += 'Currently loged in as {0}\n'.format(self._username)
        description += 'Currently slecting the {0} folder\n'.format(self._folder)
        description += 'Restricted by search rules: {0}\n'.format(self._search)
        if self._target == 'All':
            description += 'Targeting all attachments'
        else:
            if self._target:
                description += 'With following specific target\n {0}'.format([str(target) for target in self._target])
            else:
                description += 'With no target\n'
        return description
    
    def close(self):
        '''
        logout and close the imap connection
        '''
        self._imap.close()
        self._imap.logout()
        self._imap = None
    
    
    def changeFolder(self, folder, reconnect = True):
        '''
        change the folder of connection
        in case failed to connect to the new folder, 
        if reconnect is true, it will try to reselect the old folder
        '''
        result, detail = self._imap.select(folder)
        if result == 'OK':
            self._folder = folder
        else:
            print 'unable to select the {0} folder\nError: {1}'.format(folder, detail)
            if reconnect:
                result, _ = self._imap.select(self._folder)
                assert result == 'OK', 'unable to reselect old {0} folder'.format(self._folder)
                print 'reconnected to {0} folder'.format(self._folder)
        
    def search(self, keyWord = '', gmail = None):
        '''
        If gmail is true, it will use a gmail_search instead of simple search.
        This gmail_search will behave like the search in the web side.
        If gmail is false, keyWord much be a search string.
            '(FROM "Sender Name")' '(Unseen)' 'CC "CC Name")' 
            'Body "word in body"' '(Since "date only string")'
            https://tools.ietf.org/html/rfc3501.html#section-6.4.4
        If gmail is None, it will try to check if server is imap.gmail.com.
        return a list of mail id, larger id is for newer email
        '''
        self._search = keyWord
        if gmail is None:
            if self._server == 'imap.gmail.com':
                gmail = True
            else:
                gmail = False
        if gmail:
            result, emailIndex = self._imap.gmail_search(None, self._search)
        else:
            result, emailIndex = self._imap.search(None, self._search)
        assert result == 'OK', 'unable to search for {0}'.format(self._search)
        return [id for id in reversed(emailIndex[0].split())]
        
    def addTarget(self, attachment = None, subject = None, target = None, renew = False):
        '''
        update target
        if renew is false the target will be added to the existings target or initilize a new one
        if renew is true thea target will alwasy initilize to a new one
        subject is name for the email subject
        '''
        if isinstance(target, list):
            target = target
        else:
            if isinstance(attachment, str):
                target = [ {'subject': subject, 'attachment': attachment} ]
            elif isinstance(attachment, list):
                if subject == None:
                    target = [ {'subject': subject, 'attachment': att} for att in attachment]
                else:
                    target = [ {'subject': sub, 'attachment': att} for sub, att in zip(subject, attachment)]
        assert isinstance(target, list), ' Target should be All, a string or list of attachment names and target_subject should be corresponding!'
            
        if renew | (self._target == 'All'):
            self._target = target
        else:
            
            self._target.extend(target)
        
    def isTarget(self, email_attachment, email_subject):
        '''
        To judge if certain email is target or not based on its attachment or subject name
        if attavhment or subject name is None, then those part will be ignored
        '''
        if self._target == 'All':
            return True
        for tar in self._target:
            sub = tar['subject']
            att = tar['attachment']
            if ((sub == None) | (sub == email_subject)) & ((att == None) | (att == email_attachment)):
                self.removeTarget(target = {'subject': sub, 'attachment': att})
                return True    
        return False
            
    def removeTarget(self, target):
        '''
        remove a found attachment
        '''
        self._target.remove(target)
        
    def isEmptyTarget(self):
        '''
        check if target is empty or not
        '''
        return not bool(self._target)  
    
    def download(self, emailIndex, target = '', fetchLimit = 500, appendSubject = False):
        '''
        emailIndex:       iterable of all email id to be fetched
        target:         target file name list including extension
        fetchLimit:    maximum number of email to fetch
        '''
        print '******************************************\nBegin\n******************************************'
        if target:
            self._target = target
            print 'Target updated to {0}'.format(target)        
        # if attachments folder is not in the file directiory, a empty folder will be created
        baseDir = '.'
        if 'attachments' not in os.listdir(baseDir):
            os.mkdir('attachments')
            print 'made the new attachments folder'
        for fetched, index in enumerate(emailIndex):
            if fetched >= fetchLimit:
                print '******************************************\nFetch Limit {0} Reached\n'.format(fetchLimit)
                break
            if self.isEmptyTarget():
                print '******************************************\nNo More Target\n'
                break
            # the emailPackage contains email itself and some other meta data about it
            result, emailPackage = self._imap.fetch(index, '(RFC822)')
            assert result == 'OK', 'unable to fetch email with index {0}'.format(index)
            print '\nemail {0} fetched'.format(index)
            # the emailAll contains sll and different type of elements for a email
                # itself is a huge string will some filed in it
            # email.walk() will be the way to look the element in emailAll
            # email.get_content_maintype() can get content type
                # first a coupl eis usually of multipart type, which contians further subject, body and attachment, 
                # seemingly one multipart for each level of email, if the email has been forwarded or replyed,
                # it will has corresponding number of multipart
                # then the text element, which is usually the message body of the email
                # then the attachments
            # email.get('Content-Disposition') will get description for the part of email
                    # description is always None fo rmultipart and text
            emailAll = email.message_from_string(emailPackage[0][1])
            for part in emailAll.walk():
                if (part.get_content_maintype() == 'multipart') | (part.get('Content-Disposition') == None):
                    continue
                # don't be confused be the name of variable, there are possibility it is a None and not a attachment
                attachmentName = part.get_filename() # with extension
                if attachmentName:
                    attachmentName = attachmentName.replace('\r', '').replace('\n', '')
                    subjectName = emailAll['subject'].replace('\r', '').replace('\n', '')
                    if self.isTarget(email_attachment = attachmentName, email_subject = subjectName):
                        print '{0} found at email {1} ({2})\n'.format(attachmentName, index, subjectName)
                        if appendSubject:
                            newName = subjectName.replace(':', '').replace('.', '') + ' ' + attachmentName
                            print '{0} was renamed to {1}!!\n'.format(attachmentName, newName)
                            attachmentName = newName
                        filePath = os.path.join(baseDir, 'attachments', attachmentName)
                        if not os.path.isfile(filePath):
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                        else:
                            print '{0} already exist!!\n'.format(attachmentName)
                            os.remove(filePath)
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                            print '{0} was replaced!!\n'.format(attachmentName)

            print 'email {0} ({1}) processed\n'.format(index, emailAll['subject'])
        print 'End\n******************************************\n'

    @staticmethod
    def extractPhrase(urlfile, pattern = 'url'):
        '''
        this functino will extract the phrase using the regex pattern,
        url and email could be two commonn pattern, 
        otherwise, just pass the pattern string directly not the re compiled object
        '''
        assert isinstance(pattern, str), 'wrong pattern type, should be str'
        common_patterns = {'url': r'(https|http|ftp|www)(://|.)([\w.,@?![\]^=%&:/~+#-]*[\w@?![\]^=%&/~+#-])+', 
                           'email': r'([\w+_.]+@[\w-]+\.[\w.-]+)'}
        if pattern in common_patterns:
            pattern = common_patterns[pattern]
        groups = re.findall(pattern, urlfile, flags = re.MULTILINE)
        phrases = [''.join(group) for group in groups]
        return phrases
        
    def downloadPhrase(self, emailIndex, pattern = 'url', fetchLimit = 500, appendSubject = False):
        '''
        emailIndex:       iterable of all email id to be fetched
        target:         target file name list including extension
        fetchLimit:    maximum number of email to fetch
        '''
        print '******************************************\nBegin\n******************************************'
        phraseFileName = 'Phrases List'
        if os.path.isfile(os.path.join('.', phraseFileName + '.txt')):
            index = 1
            while os.path.isfile(os.path.join('.', phraseFileName + ' ' + str(index) + '.txt')):
                index += 1
            phraseFileName = phraseFileName + ' ' + str(index)
        fp = open(os.path.join('.', phraseFileName + '.txt'), 'wb')
        fp.write('*********************\r\n\r\nALL Phrases List ({0})\r\n*********************\r\n\r\n'.format(self._search))
        for fetched, index in enumerate(emailIndex):
            if fetched >= fetchLimit:
                print '******************************************\nFetch Limit {0} Reached\n'.format(fetchLimit)
                break
            # the emailPackage contains email itself and some other meta data about it
            result, emailPackage = self._imap.fetch(index, '(RFC822)')
            assert result == 'OK', 'unable to fetch email with index {0}'.format(index)
            print '\nemail {0} fetched'.format(index)
            # the emailAll contains sll and different type of elements for a email
                # itself is a huge string will some filed in it
            # email.walk() will be the way to look the element in emailAll
            # email.get_content_maintype() can get content type
                # first a coupl eis usually of multipart type, which contians further subject, body and attachment, 
                # seemingly one multipart for each level of email, if the email has been forwarded or replyed,
                # it will has corresponding number of multipart
                # then the text element, which is usually the message body of the email
                # then the attachments
            # email.get('Content-Disposition') will get description for the part of email
                    # description is always None fo rmultipart and text
            emailAll = email.message_from_string(emailPackage[0][1])
        for part in emailAll.walk():
            if part.get_content_maintype() == 'text':
                phrase_list = self.extractPhrase(part.get_payload(decode = True), pattern) # with extension
            else:
                continue
            fp.write("Phrase From {0}\r\n----------\r\n".format(emailAll['subject']))
            for phrase in phrase_list:
                fp.write("%s\r\n" % phrase)
            fp.write("----------\r\n\r\n")
            print 'email {0} ({1}) processed\n'.format(index, emailAll['subject'])
        fp.close()
        print '******************************************\nEnd\n******************************************\n'















        
        
        
        
        
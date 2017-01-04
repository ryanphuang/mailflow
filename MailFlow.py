from AppKit import NSAlternateKeyMask, NSApplication, NSBundle, NSMenuItem, \
        NSLog, NSCommandKeyMask, NSUserDefaults, NSOffState, NSOnState, NSObject
import objc
import re
import textwrap

def Category(classname):
    return objc.Category(objc.lookUpClass(classname))

def Class(classname):
    return objc.lookUpClass(classname)

def swizzle(classname, selector):
    def decorator(function):
        cls = objc.lookUpClass(classname)
        old = cls.instanceMethodForSelector_(selector)
        if old.isClassMethod:
            old = cls.methodForSelector_(selector)
        def wrapper(self, *args, **kwargs):
            return function(self, old, *args, **kwargs)
        new = objc.selector(wrapper, selector = old.selector,
                            signature = old.signature,
                            isClassMethod = old.isClassMethod)
        objc.classAddMethod(cls, selector, new)
        return wrapper
    return decorator

class DefaultsProxy:
    def __init__(self, typename, delegate):
        self.typename = typename
        self.delegate = delegate

    def __getitem__(self, item):
        return {
            'bool'   : self.delegate.boolForKey_,
            'int'    : self.delegate.integerForKey_,
            'string' : self.delegate.stringForKey_,
            'object' : self.delegate.objectForKey_,
        }[self.typename](item)

    def __setitem__(self, item, value):
        {
            'int'    : self.delegate.setInteger_forKey_,
            'bool'   : self.delegate.setBool_forKey_,
            'string' : self.delegate.setObject_forKey_,
            'object' : self.delegate.setObject_forKey_,
        }[self.typename](value, item)

class NSUserDefaults(Category('NSUserDefaults')):
    @property
    def bool(self):
        return DefaultsProxy('bool', self)

    @property
    def int(self):
        return DefaultsProxy('int', self)

    @property
    def string(self):
        return DefaultsProxy('string', self)

    @property
    def object(self):
        return DefaultsProxy('object', self)

def flow(text, width, padspace=True):
    quote, indent = re.match(r'(>+ ?|)(\s*)', text, re.UNICODE).groups()
    prefix = len(quote)
    if text[prefix:] == u'-- ':
        return [text]
    text = text.rstrip(u' ')

    if not quote:
        if indent.startswith(u' ') or text.startswith(u'From '):
            text = u' ' + text
    if indent or len(text) <= width:
        return [text]

    matches = re.finditer(r'\S+\s*(?=\S|$)', text[prefix:], re.UNICODE)
    breaks, lines = [match.end() + prefix for match in matches], []
    while True:
        for index, cursor in enumerate(breaks[1:]):
            if len(text[:cursor].expandtabs()) >= width:
                cursor = breaks[index]
                break
        else:
            lines.append(text)
            return lines
        if padspace:
            lines.append(text[:cursor] + u' ')
        else:
            lines.append(text[:cursor])
        if not quote and text[cursor:].startswith(u'From '):
            text, cursor = u' ' + text[cursor:], cursor - 1
        else:
            text, cursor = quote + text[cursor:], cursor - prefix
        breaks = [offset - cursor for offset in breaks[index + 1:]]

def wrap(text, level, width, detect_bullet_list):
    NSLog('MailFlow to wrap text')
    initial = subsequent = len(text) - len(text.lstrip())
    if detect_bullet_list and initial > 0:
        if text.lstrip().startswith(('- ', '+ ', '* ')):
            subsequent += 2
    return textwrap.fill(' '.join(text.split()),
                         width - level - 1 if level > 0 else width,
                         break_long_words = False,
                         break_on_hyphens = False,
                         initial_indent = ' ' * initial,
                         subsequent_indent = ' ' * subsequent)

class ComposeViewController(Category('ComposeViewController')):

    @classmethod
    def registerWithApplication(cls, app):
        cls.app = app

    @swizzle('ComposeViewController', '_finishLoadingEditor')
    def _finishLoadingEditor(self, old):
        result = old(self)
        if self.messageType() not in [1, 2, 3, 8]:
            return result

        view = self.composeWebView()
        document = view.mainFrame().DOMDocument()
        view.contentElement().removeStrayLinefeeds()
        blockquotes = document.getElementsByTagName_('BLOCKQUOTE')
        for index in xrange(blockquotes.length()):
            if blockquotes.item_(index):
                blockquotes.item_(index).removeStrayLinefeeds()

        if self.messageType() in [1, 2, 8]:
            if self.app.is_fix_attribution:
                view.moveToBeginningOfDocument_(None)
                view.moveToEndOfParagraphAndModifySelection_(None)
                view.moveForwardAndModifySelection_(None)
                item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
                    'Decrease', 'changeQuoteLevel:', '')
                item.setTag_(-1)
                view.changeQuoteLevel_(item)

                attribution = view.selectedDOMRange().stringValue()
                attribution = attribution.rsplit(u',', 1)[-1].lstrip()
                if view.isAutomaticTextReplacementEnabled():
                    view.setAutomaticTextReplacementEnabled_(False)
                    view.insertText_(attribution)
                    view.setAutomaticTextReplacementEnabled_(True)
                else:
                    view.insertText_(attribution)

            signature = document.getElementById_('AppleMailSignature')
            if signature:
                range = document.createRange()
                range.selectNode_(signature)
                view.setSelectedDOMRange_affinity_(range, 0)
                view.moveUp_(None)
            else:
                view.moveToEndOfDocument_(None)
                view.insertParagraphSeparator_(None)

        if self.messageType() == 3:
            for index in xrange(blockquotes.length()):
                blockquote = blockquotes.item_(index)
                if blockquote.quoteLevel() == 1:
                    blockquote.parentNode().insertBefore__(
                        document.createElement_('BR'), blockquote)

        view.insertParagraphSeparator_(None)
        view.undoManager().removeAllActions()
        self.setHasUserMadeChanges_(False)
        self.backEnd().setHasChanges_(False)
        return result

    @swizzle('ComposeViewController', 'show')
    def show(self, old):
        result = old(self)
        if self.messageType() in [1, 2, 8]:
            view = self.composeWebView()
            document = view.mainFrame().DOMDocument()
            signature = document.getElementById_('AppleMailSignature')
            if signature:
                range = document.createRange()
                range.selectNode_(signature)
                view.setSelectedDOMRange_affinity_(range, 0)
                view.moveUp_(None)
            else:
                view.moveToEndOfDocument_(None)
        return result


class EditingMessageWebView(Category('EditingMessageWebView')):

    @classmethod
    def registerWithApplication(cls, app):
        cls.app = app

    @swizzle('EditingMessageWebView', 'decreaseIndentation:')
    def decreaseIndentation_(self, original, sender, indent = 2):
        if self.contentElement().className() != 'ApplePlainTextBody':
            return original(self, sender)

        self.undoManager().beginUndoGrouping()
        affinity = self.selectionAffinity()
        selection = self.selectedDOMRange()

        self.moveToBeginningOfParagraph_(None)
        if selection.collapsed():
            for _ in xrange(indent):
                self.moveForwardAndModifySelection_(None)
            text = self.selectedDOMRange().stringValue() or ''
            if re.match(u'[ \xa0]{%d}' % indent, text, re.UNICODE):
                self.deleteBackward_(None)
        else:
            while selection.compareBoundaryPoints__(1, # START_TO_END
                    self.selectedDOMRange()) > 0:
                for _ in xrange(indent):
                    self.moveForwardAndModifySelection_(None)
                text = self.selectedDOMRange().stringValue() or ''
                if re.match(u'[ \xa0]{%d}' % indent, text, re.UNICODE):
                    self.deleteBackward_(None)
                else:
                    self.moveBackward_(None)
                self.moveToEndOfParagraph_(None)
                self.moveForward_(None)

        self.setSelectedDOMRange_affinity_(selection, affinity)
        self.undoManager().endUndoGrouping()

    @swizzle('EditingMessageWebView', 'increaseIndentation:')
    def increaseIndentation_(self, original, sender, indent = 2):
        if self.contentElement().className() != 'ApplePlainTextBody':
            return original(self, sender)

        self.undoManager().beginUndoGrouping()
        affinity = self.selectionAffinity()
        selection = self.selectedDOMRange()

        if selection.collapsed():
            position = self.selectedRange().location
            self.moveToBeginningOfParagraph_(None)
            position -= self.selectedRange().location
            self.insertText_(indent * u' ')
            for _ in xrange(position):
                self.moveForward_(None)
        else:
            self.moveToBeginningOfParagraph_(None)
            while selection.compareBoundaryPoints__(1, # START_TO_END
                    self.selectedDOMRange()) > 0:
                self.moveToEndOfParagraphAndModifySelection_(None)
                if not self.selectedDOMRange().collapsed():
                    self.moveToBeginningOfParagraph_(None)
                    self.insertText_(indent * u' ')
                    self.moveToEndOfParagraph_(None)
                self.moveForward_(None)
            self.setSelectedDOMRange_affinity_(selection, affinity)

        self.undoManager().endUndoGrouping()

    def wrapParagraph(self):
        # Note the quote level of the current paragraph and the location of
        # the end of the message to avoid attempts to move beyond it.

        NSLog('MailFlow wrap paragraph')

        self.moveToEndOfDocumentAndModifySelection_(None)
        last = self.selectedRange().location + self.selectedRange().length

        self.moveToBeginningOfParagraph_(None)
        self.selectParagraph_(None)
        level = self.quoteLevelAtStartOfSelection()

        # If we are on a blank line, move down to the start of the next
        # paragraph block and finish.

        if not self.selectedText().strip():
            while True:
                self.moveDown_(None)
                self.selectParagraph_(None)
                location = self.selectedRange().location
                if location + self.selectedRange().length >= last:
                    self.moveToEndOfParagraph_(None)
                    NSLog('MailFlow wrap paragraph done 1')
                    return
                if self.selectedText().strip():
                    self.moveToBeginningOfParagraph_(None)
                    NSLog('MailFlow wrap paragraph done 2')
                    return

        # Otherwise move to the start of this paragraph block, working
        # upward until we hit the start of the message, a blank line or a
        # change in quote level.

        NSLog('MailFlow paragraph start')
        while self.selectedRange().location > 0:
            self.moveUp_(None)
            if self.quoteLevelAtStartOfSelection() != level:
                self.moveDown_(None)
                break
            self.selectParagraph_(None)
            if not self.selectedText().strip():
                self.moveDown_(None)
                break
        self.moveToBeginningOfParagraph_(None)

        # Insert a temporary placeholder space character to avoid any
        # assumptions about Mail.app's strange and somewhat unpredictable
        # handling of newlines between block elements.

        self.insertText_(' ')
        self.moveToEndOfParagraphAndModifySelection_(None)

        # Now extend the selection forward line-by-line until we hit a blank
        # line, a change in quote level or the end of the message.

        affinity = self.selectionAffinity()
        selection = self.selectedDOMRange()
        while True:
            location = self.selectedRange().location
            if location + self.selectedRange().length >= last:
                break
            self.moveDown_(None)
            self.moveToEndOfParagraphAndModifySelection_(None)
            if self.quoteLevelAtStartOfSelection() != level:
                break
            if not self.selectedText().strip():
                break
            selection.setEnd__(self.selectedDOMRange().endContainer(),
                               self.selectedDOMRange().endOffset())
        self.setSelectedDOMRange_affinity_(selection, affinity)

        # Finally, extend the selection forward to encompass any blank lines
        # following the paragraph block, regardless of quote level. Store
        # the minimum quote level of this paragraph block and the next.

        while True:
            location = self.selectedRange().location
            if location + self.selectedRange().length >= last:
                minimum = 0
                break
            self.moveDown_(None)
            self.moveToBeginningOfParagraph_(None)
            self.moveToEndOfParagraphAndModifySelection_(None)
            if self.selectedText().strip():
                minimum = min(self.quoteLevelAtStartOfSelection(), level)
                break
            selection.setEnd__(self.selectedDOMRange().endContainer(),
                               self.selectedDOMRange().endOffset())
        self.setSelectedDOMRange_affinity_(selection, affinity)
        NSLog('MailFlow scanned paragraph: wrap width=%d, detect_bullet_list=%s' % (self.app.wrap_width, self.app.detect_bullet_list))

        # Re-fill the text allowing for quote level and retaining block
        # indentation, then insert it to replace the selection.

        NSLog('MailFlow before wrap')
        text = wrap(self.selectedText().expandtabs(), level, self.app.wrap_width, self.app.detect_bullet_list) + '\n'
        NSLog('MailFlow done wrap')
        self.insertTextWithoutReplacement_(text)

        # Reduce the quote level of the trailing blank line if necessary,
        # then remove the placeholder character and position the cursor at
        # the start of the next paragraph block.

        item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
            'Decrease', 'changeQuoteLevel:', '')
        item.setTag_(-1)
        for _ in xrange(level - minimum):
            self.changeQuoteLevel_(item)

        selection = self.selectedDOMRange()
        for _ in xrange(text.count('\n')):
            self.moveUp_(None)
            self.moveToBeginningOfParagraph_(None)
        self.deleteForward_(None)
        self.setSelectedDOMRange_affinity_(selection, affinity)
        self.moveForward_(None)

    def flowText_(self, sender):
        self.app.is_flow_text = not sender.state()
        self.app.menu.flow_menu_item.setState_(self.app.is_flow_text)
        if self.app.is_flow_text:
            self.app.is_wrap_text = False
            self.app.menu.wrap_menu_item.setState_(NSOffState)

    def wrapText_(self, sender):
        self.app.is_wrap_text = not sender.state()
        self.app.menu.wrap_menu_item.setState_(self.app.is_wrap_text)
        if self.app.is_wrap_text:
            self.app.is_flow_text = False
            self.app.menu.flow_menu_item.setState_(NSOffState)

    def wrapOnce_(self, sender):
        self.app.menu.wrap_once_menu_item.setState_(sender.state())
        # Wrap text only works correctly on plain text messages, so ignore
        # any requests to format paragraphs in rich-text/HTML messages.

        if self.contentElement().className() != 'ApplePlainTextBody':
            return

        # If we have a selection, format all paragraph blocks which overlap
        # it. Otherwise, format the paragraph block containing the cursor.
        # Combine the operation into a single undo group for UI purposes.

        self.undoManager().beginUndoGrouping()
        if self.selectedRange().length == 0:
            # self.selectAll_(None)
            self.wrapParagraph()
            self.moveToEndOfDocumentAndModifySelection_(None)
        else:
            last = self.selectedRange().length
            self.moveToEndOfDocumentAndModifySelection_(None)
            last = self.selectedRange().length - last
            while self.selectedRange().length > last:
                self.wrapParagraph()
                self.moveToEndOfDocumentAndModifySelection_(None)
        if self.selectedRange().length > 0:
            self.moveBackward_(None)
        else:
            self.deleteBackward_(None)
        self.undoManager().endUndoGrouping()

    def insertTextWithoutReplacement_(self, text):
        if self.isAutomaticTextReplacementEnabled():
            self.setAutomaticTextReplacementEnabled_(False)
            self.insertText_(text)
            self.setAutomaticTextReplacementEnabled_(True)
        else:
            self.insertText_(text)

    def quoteLevelAtStartOfSelection(self):
        return self.selectedDOMRange().startContainer().quoteLevel()

    def selectedText(self):
        return self.selectedDOMRange().stringValue() or ''

class MCMessage(Category('MCMessage')):
    @swizzle('MCMessage', 'forwardedMessagePrefixWithSpacer:')
    def forwardedMessagePrefixWithSpacer_(self, old, *args):
        return u''

class MCMessageGenerator(Category('MCMessageGenerator')):

    @classmethod
    def registerWithApplication(cls, app):
        cls.app = app

    @swizzle('MCMessageGenerator', '_encodeDataForMimePart:withPartData:')
    def _encodeDataForMimePart_withPartData_(self, old, part, data):
        if part.type() != 'text' or part.subtype() != 'plain':
            return old(self, part, data)

        text = bytes(data.objectForKey_(part))
        if any(len(line) > 998 for line in text.splitlines()):
            return old(self, part, data)

        try:
            text.decode('ascii')
            part.setContentTransferEncoding_('7bit')
        except UnicodeDecodeError:
            part.setContentTransferEncoding_('8bit')
        return True

    @swizzle('MCMessageGenerator',
             '_newPlainTextPartWithAttributedString:partData:')

    def _newPlainTextPartWithAttributedString_partData_(self, old, *args):
        if not self.app.should_wrap:
            return old(self, *args)
        event = NSApplication.sharedApplication().currentEvent()
        result = old(self, *args)
        if event and event.modifierFlags() & NSAlternateKeyMask:
            return result

        charset = result.bodyParameterForKey_('charset') or 'utf-8'
        data = args[1].objectForKey_(result)
        lines = bytes(data).decode(charset).split('\n')
        lines = [line for text in lines for line in flow(text, self.app.wrap_width + 1)]
        data.setData_(buffer(u'\n'.join(lines).encode(charset)))

        result.setBodyParameter_forKey_('yes', 'delsp')
        if self.app.is_flow_text:
            result.setBodyParameter_forKey_('flowed', 'format')
        return result


class MCMimePart(Category('MCMimePart')):
    @swizzle('MCMimePart', '_decodeTextPlain')
    def _decodeTextPlain(self, old):
        result = old(self)
        if result.startswith(u' '):
            result = u'&nbsp;' + result[1:]
        return result.replace(u'<BR> ', u'<BR>&nbsp;')


class MessageViewController(Category('MessageViewController')):
    @classmethod
    def registerWithApplication(cls, app):
        cls.app = app

    @swizzle('MessageViewController', 'forward:')
    def forward_(self, old, *args):
        if not self.app.should_wrap:
            return old(self, *args)
        event = NSApplication.sharedApplication().currentEvent()
        if event and event.modifierFlags() & NSAlternateKeyMask:
            return old(self, *args)
        return self._messageViewer().forwardAsAttachment_(*args)

class MessageViewer(Category('MessageViewer')):

    @classmethod
    def registerWithApplication(cls, app):
        cls.app = app

    @swizzle('MessageViewer', 'forwardMessage:')
    def forwardMessage_(self, old, *args):
        if not self.app.should_wrap:
            return old(self, *args)
        event = NSApplication.sharedApplication().currentEvent()
        if event and event.modifierFlags() & NSAlternateKeyMask:
            return old(self, *args)
        return self.forwardAsAttachment_(*args)

class SingleMessageViewer(Category('SingleMessageViewer')):
    @classmethod
    def registerWithApplication(cls, app):
        cls.app = app

    @swizzle('SingleMessageViewer', 'forwardMessage:')
    def forwardMessage_(self, old, *args):
        if not self.app.should_wrap:
            return old(self, *args)
        event = NSApplication.sharedApplication().currentEvent()
        if event and event.modifierFlags() & NSAlternateKeyMask:
            return old(self, *args)
        return self.forwardAsAttachment_(*args)

class App(object):
    def __init__(self, version):
        self.version = version
        self.prefs = NSUserDefaults.standardUserDefaults()
        self.prefs.registerDefaults_(dict(
            FlowText = False,
            WrapText = False,
            WrapOnce = False,
            FixAttribution = False,
            BulletLists = True,
            IndentWidth = 2,
            WrapWidth = 76,
        ))
        self.menu = MailFlowMenu.alloc().initWithApp_(self).inject()

    @property
    def should_wrap(self):
        return self.is_flow_text or self.is_wrap_text

    @property
    def is_flow_text(self):
        return self.prefs.bool["FlowText"]

    @is_flow_text.setter
    def is_flow_text(self, value):
        self.prefs.bool["FlowText"] = value

    @property
    def is_wrap_text(self):
        return self.prefs.bool["WrapText"]

    @is_wrap_text.setter
    def is_wrap_text(self, value):
        self.prefs.bool["WrapText"] = value

    @property
    def is_fix_attribution(self):
        return self.prefs.bool["FixAttribution"]

    @is_fix_attribution.setter
    def is_fix_attribution(self, value):
        self.prefs.bool["FixAttribution"] = value

    @property
    def detect_bullet_list(self):
        return self.prefs.bool["BulletLists"]

    @detect_bullet_list.setter
    def detect_bullet_list(self, value):
        self.prefs.bool["BulletLists"] = value

    @property
    def indent_width(self):
        return self.prefs.int["IndentWidth"]

    @indent_width.setter
    def indent_width(self, value):
        self.prefs.int["IndentWidth"] = value

    @property
    def wrap_width(self):
        return self.prefs.int["WrapWidth"]

    @wrap_width.setter
    def wrap_width(self, value):
        self.prefs.int["WrapWidth"] = value

class MailFlowMenu(NSObject):
    def initWithApp_(self, app):
        self = objc.super(MailFlowMenu, self).init()
        if self is None:
            return None
        self.app = app
        self.mainwindow = NSApplication.sharedApplication().mainWindow()
        self.bundle = NSBundle.bundleWithIdentifier_('uk.me.cdw.MailFlow')
        return self

    def inject(self):
        NSLog('Trying to inject menu in MailFlow')
        self.retain()
        application = NSApplication.sharedApplication()
        editmenu = application.mainMenu().itemAtIndex_(2).submenu()
        editmenu.addItem_(NSMenuItem.separatorItem())

        mask = NSCommandKeyMask | NSAlternateKeyMask
        self.flow_menu_item = editmenu.addItemWithTitle_action_keyEquivalent_(
            "Flow Text",
            "flowText:", 
            "=")
        self.flow_menu_item.setKeyEquivalentModifierMask_(mask)
        self.flow_menu_item.setToolTip_("Send flow format plain-text email")
        self.flow_menu_item.setState_(self.app.is_flow_text)

        mask = NSCommandKeyMask | NSAlternateKeyMask
        self.wrap_menu_item = editmenu.addItemWithTitle_action_keyEquivalent_(
            "Wrap Text",
            "wrapText:", 
            "\\")
        self.wrap_menu_item.setKeyEquivalentModifierMask_(mask)
        self.wrap_menu_item.setToolTip_("Wrap plain-text email")
        self.wrap_menu_item.setState_(self.app.is_wrap_text)

        mask = NSCommandKeyMask
        self.wrap_once_menu_item = editmenu.addItemWithTitle_action_keyEquivalent_(
            "Wrap Once",
            "wrapOnce:", 
            "\\")
        self.wrap_once_menu_item.setKeyEquivalentModifierMask_(mask)
        self.wrap_once_menu_item.setToolTip_("Wrap selected/all text once")
        return self

class MailFlow(Class('MVMailBundle')):
    @classmethod
    def initialize(self):
        self.registerBundle()

        bundle = NSBundle.bundleWithIdentifier_('uk.me.cdw.MailFlow')
        version = bundle.infoDictionary().get('CFBundleVersion', '??')

        app = App(version)

        MCMessageGenerator.registerWithApplication(app)
        MessageViewController.registerWithApplication(app)
        MessageViewer.registerWithApplication(app)
        SingleMessageViewer.registerWithApplication(app)
        EditingMessageWebView.registerWithApplication(app)
        ComposeViewController.registerWithApplication(app)

        NSLog('Loaded MailFlow')

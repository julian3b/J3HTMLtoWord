var $jq = jQuery.noConflict();

var J3 = J3 || {};
J3.HTMLtoWord = function () {
    var  GenerateWordFile = function (params, done) {
            if (params == null) {
                done(false);
            }
            var MIME_Head = 'MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary="----=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI"\r\n\r\n',
                MIME_Document_Head = 'Content-Location: file:///C:/document.htm\r\n' +
                    'Content-Transfer-Encoding: base64\r\n' +
                    'Content-Type: text/html; charset="utf-8"\r\n\r\n',
                MIME_Header_Head = 'Content-Location: file:///C:/document_files/headerfooter.htm\r\n' +
                    'Content-Transfer-Encoding: base64\r\n' +
                    'Content-Type: text/html; charset="utf-8"\r\n\r\n',
                MIME_Header_Head2 = 'Content-Location: file:///C:/document_files/headerfooter2.htm\r\n' +
                    'Content-Transfer-Encoding: base64\r\n' +
                    'Content-Type: text/html; charset="utf-8"\r\n\r\n',
                MIME_Header_TESTAttachment = 'Content-Location: file:///C:/document_files/test.txt\r\n' +
                    'Content-Transfer-Encoding: base64\r\n' +
                    'Content-Type: text/html; charset="utf-8"\r\n\r\n', Boundary = '------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI\r\n',
                lastBoundary = '------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI--', FileDoc = '', FileHead = '';


            params = setParameters(params);

            var BodyTemplate = generateTemplateBody(params);
            var HeadFooterTemplates = generateTemplateHeadFooter(params);
            var HeadFooterTemplate1 = HeadFooterTemplates[0];
            var HeadFooterTemplate2 = HeadFooterTemplates.length > 1 ? HeadFooterTemplates[1] : '';
            imageBodyParser(BodyTemplate, function (returnBody) {
                var FileDoc = returnBody.replace(/{{PageNum}}/g, generatePagingTemplate(params)).replace(/{{TOC}}/g, generateTableOfContents(params));
                imageBodyParser(HeadFooterTemplate1, function (returnHead1) {
                var FileHead = returnHead1.replace(/{{PageNum}}/g, generatePagingTemplate(params));
                    imageBodyParser(HeadFooterTemplate2, function (returnHead2) {
                        var FileHead2 = returnHead2.replace(/{{PageNum}}/g, generatePagingTemplate(params));
                        var ReplaceFootNotes = addFootNotes(FileDoc, params.FootNotes);
                        FileDoc = ReplaceFootNotes.Body.replace(/{{FootNotes}}/g, ReplaceFootNotes.Footnotes);
                        var FileDocBase64 = btoa(unescape(encodeURIComponent(FileDoc))) + '\r\n\r\n';
                        var FileHeadBase64 = btoa(unescape(encodeURIComponent(FileHead))) + '\r\n\r\n';
                        var FileHead2Base64 = FileHead2 != '' ? btoa(unescape(encodeURIComponent(FileHead2))) + '\r\n\r\n' : '';
                        var testFile = btoa(unescape(encodeURIComponent('asdasdasdasdasdasf as fsa df'))) + '\r\n\r\n';
                        var str = MIME_Head + Boundary + MIME_Document_Head + FileDocBase64 + Boundary + MIME_Header_Head + FileHeadBase64 + Boundary + 
                            (FileHead2 != '' ? MIME_Header_Head2 : '') + (FileHead2 != '' ? FileHead2Base64 : '') + (FileHead2 != '' ? Boundary : '') +
                            MIME_Header_TESTAttachment + testFile + lastBoundary;
                        var FileDocBlob = new Blob([str], { type: "application/msword;charset=utf-8", });
                        if (params.Download) 
                            saveAs(FileDocBlob, params.FileName + ".doc");
                        done(FileDocBlob);
                    });
                });
            });
        },
        setParameters = function (params) {
            var parameters = {};
            parameters.FileName= params.FileName == null ? 'Document' : params.FileName;
            parameters.Title = params.Title == null ? '' : params.Title;
            parameters.Orientation = params.Orientation == null ? 'portrait' : params.Orientation;
            parameters.Size = params.Size == null ? '8.5in 11in' : params.Size;
            parameters.Margin= params.Margin == null ? '1in 1in 1in 1in' : params.Margin;
            parameters.HeaderMargin= params.HeaderMargin == null ? '.1in' : params.HeaderMargin;
            parameters.FooterMargin= params.FooterMargin == null ? '.5in' : params.FooterMargin;
            parameters.PageNumberText1= params.PageNumberText1 == null ? 'Page' : params.PageNumberText1;
            parameters.PageNumberText2= params.PageNumberText2 == null ? 'of' : params.PageNumberText2;
            parameters.PageNumberText3= params.PageNumberText3 == null ? '' : params.PageNumberText3;
            parameters.TOCTitle= params.TOCTitle == null ? 'Table of Contents' : params.TOCTitle;
            parameters.Download= params.Download == null ? true : params.Download;
            parameters.HeaderFooters = [];
            parameters.Pages = [];
            parameters.FootNotes = [];

            parameters.Pages= ((params.Pages == null) ? [''] : ((params.Pages == '') ? [''] : params.Pages));
            if(params.HeaderFooters == null){
                parameters.HeaderFooters =  [];
            }
            else {
                if(parameters.HeaderFooters.length > 0){
                    for(var i = 0 ; i < parameters.HeaderFooters.length; i++){
                        parameters.HeaderFooters.push([{
                            Header: parameters.HeaderFooters[i].Header != null ? parameters.HeaderFooters[i].Header : '', 
                            Footer: parameters.HeaderFooters[i].Footer != null ? parameters.HeaderFooters[i].Footer : ''}]);
                    }
                }
                else
                    parameters.HeaderFooters =  [];
            }
            parameters.FootNotes=((params.FootNotes == null) ? [] : ((params.FootNotes == '') ? [] : params.FootNotes));
            return parameters;
        },
        imageBodyParser = function (body, callback) {
            var newBody = $jq(body);
            var totalImg = $jq(body).find("img").length;
            if (totalImg > 0) {
                var totalConverted = 0;
                $jq(newBody).find("img").each(function (index, e) {
                    var imgBase64 = '';
                    var width = e.width;
                    var height = e.height;

                    image_to_base64(e.src, width, height, function (result) {
                        imgBase64 = result;
                        e.src = imgBase64;
                        totalConverted++;
                        if (totalConverted == totalImg) {
                            var returnText = '';
                            for(var i = 0 ; i < newBody.length ; i++){
                                if(newBody[i].outerHTML != null)
                                    returnText += newBody[i].outerHTML;
                                else if(newBody[i].nodeType== 3)
                                    returnText += newBody[i].textContent;
                                      
                            }
                            callback(returnText);
                        }
                    });
                });
            }
            else
                callback(body);
        },     
        image_to_base64 = function (file, fwidth, fheight, callback) {
            var httpRequest = new XMLHttpRequest();
            httpRequest.onload = function () {
                var fileReader = new FileReader();

                fileReader.onloadend = function () {

                    if(fwidth == 0 &&  fheight == 0){
                        callback(fileReader.result);
                    }
                    else{
                        var img = new Image();
                        img.src = fileReader.result;
                        img.onload = function() {
                        
                        var newWidth;
                        var newHeight;
                        if(fwidth > 0 &&  fheight == 0){
                            newWidth = calculateAspectRatioFit(this.width, this.height, fwidth, this.height).width;
                            newHeight = calculateAspectRatioFit(this.width, this.height, fwidth, this.height).height;
                        }
                        else if(fwidth == 0 &&  fheight > 0){
                            newWidth = calculateAspectRatioFit(this.width, this.height, this.width, fheight).width;
                            newHeight = calculateAspectRatioFit(this.width, this.height, this.width, fheight).height;
                        }
                        else{
                            newWidth = fwidth;
                            newHeight = fheight;
                        }

                        var img = document.createElement("img");
                        img.src = fileReader.result;
                        img.width = newWidth;
                        img.height = newHeight;
    
                        var canvas = document.createElement("canvas");
                        var ctx = canvas.getContext("2d");
                        ctx.drawImage(img, 0, 0);
    
                        canvas.width = newWidth;
                        canvas.height = newHeight;
                        var ctx = canvas.getContext("2d");
                        ctx.drawImage(img, 0, 0, newWidth, newHeight);
            
                        var dataurl = canvas.toDataURL("image/png");
    
                        callback(dataurl);
                    }
                    }

                }
                fileReader.readAsDataURL(httpRequest.response);
            };
            httpRequest.open('GET', file);
            httpRequest.responseType = 'blob';
            httpRequest.send();
        },
        generateTemplateBody = function(DocParams){
            var TemplateDoc = '';
            TemplateDoc += '<head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><title>' + DocParams.Title + '</title>' +
            '<link rel=File-List href="document_files/filelist.xml">' +
            '    <style>' +
            '        v\:* {behavior:url(#default#VML);}' +
            '        o\:* {behavior:url(#default#VML);}' +
            '        w\:* {behavior:url(#default#VML);}' +
            '        .shape {behavior:url(#default#VML);}' +
            '        </style>' +
            '        <style>' +
            '            p.MsoHeader, li.MsoHeader, div.MsoHeader{' +
            '                    margin:0in;' +
            '                    margin-top:.0001pt;' +
            '                    mso-pagination:widow-orphan;' +
            '                    tab-stops:center 3.0in right 6.0in;' +
            '                }' +
            '                p.MsoFooter, li.MsoFooter, div.MsoFooter{' +
            '                    margin:0in;' +
            '                    margin-bottom:.0001pt;' +
            '                    mso-pagination:widow-orphan;' +
            '                    tab-stops:center 3.0in right 6.0in;' +
            '                }' +
            '            p.MsoTocHeading, li.MsoTocHeading, div.MsoTocHeading'+
            '    {mso-style-priority:39;'+
            '    mso-style-qformat:yes;'+
            '    mso-style-parent:"Heading 1";'+
            '    mso-style-next:Normal;'+
            '    margin-top:12.0pt;'+
            '    margin-right:0in;'+
            '    margin-bottom:0in;'+
            '    margin-left:0in;'+
            '    margin-bottom:.0001pt;'+
            '    line-height:107%;'+
            '    mso-pagination:widow-orphan lines-together;'+
            '    page-break-after:avoid;'+
            '    font-size:16.0pt;'+
            '    font-family:"Calibri Light",sans-serif;'+
            '    mso-ascii-font-family:"Calibri Light";'+
            '    mso-ascii-theme-font:major-latin;'+
            '    mso-fareast-font-family:"Times New Roman";'+
            '    mso-fareast-theme-font:major-fareast;'+
            '    mso-hansi-font-family:"Calibri Light";'+
            '    mso-hansi-theme-font:major-latin;'+
            '    mso-bidi-font-family:"Times New Roman";'+
            '    mso-bidi-theme-font:major-bidi;'+
            '    color:#2F5496;'+
            '    mso-themecolor:accent1;'+
            '    mso-themeshade:191;}'+
            'p.MsoFootnoteText, li.MsoFootnoteText, div.MsoFootnoteText'+
            '{  mso-style-noshow:yes;'+
            '   mso-style-priority:99;'+
            '   so-style-link:"Footnote Text Char";'+
            '   margin:0in;'+
            '   margin-bottom:.0001pt;'+
            '   mso-pagination:widow-orphan;'+
            '   font-size:10.0pt;'+
            '   font-family:"Calibri",sans-serif;'+
            '   mso-ascii-font-family:Calibri;'+
            '   mso-ascii-theme-font:minor-latin;'+
            '   mso-fareast-font-family:Calibri;'+
            '   mso-fareast-theme-font:minor-latin;'+
            '   mso-hansi-font-family:Calibri;'+
            '   mso-hansi-theme-font:minor-latin;'+
            '   mso-bidi-font-family:"Times New Roman";'+
            '   mso-bidi-theme-font:minor-bidi;}'+
            'span.MsoFootnoteReference'+
            '{  mso-style-noshow:yes;'+
            '    mso-style-priority:99;'+
            '    vertical-align:super;}'+
            'span.FootnoteTextChar'+
            '{  mso-style-name:"Footnote Text Char";'+
            '   mso-style-noshow:yes;'+
            '   mso-style-priority:99;'+
            '   mso-style-unhide:no;'+
            '   mso-style-locked:yes;'+
            '   mso-style-link:"Footnote Text";'+
            '   mso-ansi-font-size:10.0pt;'+
            '   mso-bidi-font-size:10.0pt;} '+           
            '' +
            '        @page' +
            '        {' +
            '            mso-page-orientation: ' + DocParams.Orientation + ';' +
            '            size:' + DocParams.Size + ';    margin:' + DocParams.Margin + ';' +
            '            mso-footnote-separator:url("document_files/headerfooter.htm") fs;' +
            '	         mso-footnote-continuation-separator:url("document_files/headerfooter.htm") fcs;' +
            '	         mso-endnote-separator:url("document_files/headerfooter.htm") es;' +
            '	         mso-endnote-continuation-separator:url("document_files/headerfooter.htm") ecs;}' +
            '        }' +
            '        @page Section1 {' +
            '            mso-header-margin:' + DocParams.HeaderMargin + ';' +
            '            mso-footer-margin:' + DocParams.FooterMargin + ';' +
            '            mso-title-page: yes;';

            if(DocParams.HeaderFooters.length == 2){
                TemplateDoc += ' mso-first-header: url("document_files/headerfooter.htm") h1;' +
                               ' mso-first-footer: url("document_files/headerfooter.htm") f1;' +
                               ' mso-header: url("document_files/headerfooter2.htm") h2;' +
                               ' mso-footer: url("document_files/headerfooter2.htm") f2;';
            }
            else{
                TemplateDoc +='mso-first-header: url("document_files/headerfooter.htm") h1;' +
                'mso-first-footer: url("document_files/headerfooter.htm") f1;';
            }

            TemplateDoc +=' mso-paper-source:0;' +
                          ' }' +
                          ' div.Section1 {' +
                          ' page: Section1;' +
                          ' }' +
                          ' </style>' +
                          ' <xml>' +
                          '        <word:WordDocument>' +
                          '            <word:View>Print</word:View>' +
                          '            <word:Zoom>90</word:Zoom>' +
                          '            <word:DoNotOptimizeForBrowser />' +
                          '        </word:WordDocument>' +
                          ' </xml>' +
                          '</head>' +
                          '<body>' +
                          '    <div class="Section1">';
            for(var i = 0 ; i < DocParams.Pages.length ; i++){
                TemplateDoc += DocParams.Pages[i]
                if(i < (DocParams.Pages.length - 1))
                {
                    TemplateDoc += '<br clear="all" style="page-break-before:always" />';
                }
            }
            TemplateDoc +='    </div>{{FootNotes}}</body>';

            return TemplateDoc;
        },
        generateTemplateHeadFooter = function(DocParams){
            var FootNoteTemplateSeparator = '<div style=\'mso-element:footnote-separator\' id=fs>' +
                                            '                          <p class=MsoNormal><span style=\'mso-special-character:footnote-separator\'><![if !supportFootnotes]>' +
                                            '                          <hr align=left size=1 width="33%">' +
                                            '                          <![endif]></span></p>' +
                                            '                       </div>' +
                                            '                       <div style=\'mso-element:footnote-continuation-separator\' id=fcs>' +
                                            '                          <p class=MsoNormal><span style=\'mso-special-character:footnote-continuation-separator\'><![if !supportFootnotes]>' +
                                            '                          <hr align=left size=1>' +
                                            '                          <![endif]></span></p>' +
                                            '                       </div>' +
                                            '                       <div style=\'mso-element:endnote-separator\' id=es>' +
                                            '                         <p class=MsoNormal><span style=\'mso-special-character:footnote-separator\'><![if !supportFootnotes]>' +
                                            '                         <hr align=left size=1 width="33%">' +
                                            '                         <![endif]></span></p>' +
                                            '                       </div>' +
                                            '                       <div style=\'mso-element:endnote-continuation-separator\' id=ecs>' +
                                            '                          <p class=MsoNormal><span style=\'mso-special-character:footnote-continuation-separator\'><![if !supportFootnotes]>' +
                                            '                          <hr align=left size=1>' +
                                            '                          <![endif]></span></p>' +
                                            '                       </div>';

            var templates = [];
            if(DocParams.HeaderFooters.length == 0){
                templates.push('<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"= xmlns="http://www.w3.org/TR/REC-html40">' +
                    '<body>' +
                    FootNoteTemplateSeparator +
                    '<div style="mso-element:header;" id="h1">' +
                    '<p class=MsoHeader></p>' +
                    '</div>' +
                    '<div style=\'mso-element:footer\' id=f1>' +
                    '<p class=MsoFooter></p>' +
                    '</div></body></html>');
            }
            else{
                for(var i = 0 ; i < DocParams.HeaderFooters.length ; i++){
                    templates.push('<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"= xmlns="http://www.w3.org/TR/REC-html40">' +
                        '<body>' +
                        FootNoteTemplateSeparator +
                        '<div style="mso-element:header;" id="h' + (i + 1) + '">' +
                        '<p class=MsoHeader>' + (DocParams.HeaderFooters[i].Header != null ? DocParams.HeaderFooters[i].Header : '' ) + '</p>' +
                        '</div>' +
                        '<div style=\'mso-element:footer\' id=f' + (i + 1) + '>' +
                        '<p class=MsoFooter>' + (DocParams.HeaderFooters[i].Footer != null ? DocParams.HeaderFooters[i].Footer : '' ) + '</p>' +
                        '</div></body></html>');
                }
            }
            return templates;
        },
        generatePagingTemplate = function(parameters, pagerNum){

            return PagingTemplate = '<span class=SpellE></span> ' + (parameters.PageNumberText1 == null ? "Page" : parameters.PageNumberText1 ) + ' <!--[if supportFields]><span' +
            '             class=MsoPageNumber><span style=\'mso-element:field-begin\'></span><span' +
            '             style=\'mso-spacerun:yes\'> </span>PAGE <span style=\'mso-element:field-separator\'></span></span><![endif]--><span' +
            '             class=MsoPageNumber><span style=\'mso-no-proof:yes\'>1</span></span><!--[if supportFields]><span' +
            '             class=MsoPageNumber><span style=\'mso-element:field-end\'></span></span><![endif]--><span' +
            '             class=MsoPageNumber> ' + (parameters.PageNumberText2 == null ? "of" : parameters.PageNumberText2 ) + ' </span><!--[if supportFields]><span class=MsoPageNumber><span' +
            '             style=\'mso-element:field-begin\'></span> NUMPAGES <span style=\'mso-element:field-separator\'></span></span><![endif]--><span' +
            '             class=MsoPageNumber><span style=\'mso-no-proof:yes\'>1</span></span><!--[if supportFields]><span' +
            '             class=MsoPageNumber><span style=\'mso-element:field-end\'></span></span>' + (parameters.PageNumberText3 == null ? "of" : parameters.PageNumberText3 ) + '<![endif]-->';

        },
        generateTableOfContents = function(parameters){
            return '<w:Sdt SdtDocPart="t" DocPartType="Table of Contents" DocPartUnique="t"'+
            ' ID="489288532">'+
            ' <p class=MsoTocHeading>' + parameters.TOCTitle + '<w:sdtPr></w:sdtPr></p>'+
            ' <p class=MsoNormal><!--[if supportFields]><span style=\'mso-element:field-begin\'></span><span'+
            ' style=\'mso-spacerun:yes\'></span> TOC \\o &quot;1-3&quot; \\h \\z \\u <span'+
            ' style=\'mso-element:field-separator\'></span><![endif]--><b><span'+
            ' style=\'mso-no-proof:yes\'>Please select the "Update Table" button to refresh the table.</span></b><!--[if supportFields]><b><span'+
            ' style=\'mso-no-proof:yes\'><span style=\'mso-element:field-end\'></span></span></b><![endif]--></p>'+
           ' </w:Sdt>';
        
        },
        calculateAspectRatioFit = function(srcWidth, srcHeight, maxWidth, maxHeight) {
            var ratio = Math.min(maxWidth / srcWidth, maxHeight / srcHeight);
            return { width: srcWidth*ratio, height: srcHeight*ratio };
        },
        addFootNotes = function(Body, FootNotesArray){
            if(FootNotesArray.length > 0){
                var BodyReturn = Body;
                var FootNotesString = '<div style=\'mso-element:footnote-list\'><![if !supportFootnotes]><br clear=all>'+
                '<hr align=left size=1 width="33%">'+
                '<![endif]>';

                for(var i = 0 ; i < FootNotesArray.length ; i++){
                var footnoteReference = '<a style=\'mso-footnote-id:ftn' + (i + 1) + '\' href="#_ftn' + (i + 1) + '" name="_ftnref' + (i + 1) + '" '+
                ' title=""><span class=MsoFootnoteReference><span style=\'mso-special-character:'+
                ' footnote\'><![if !supportFootnotes]><span class=MsoFootnoteReference>'+
                '<span style=\'font-size:12.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:'+
                '"Times New Roman";mso-ansi-language:EN-US;mso-fareast-language:EN-US;'+
                'mso-bidi-language:AR-SA\'>[' + (i + 1) + ']</span></span><![endif]></span></span></a>';

                    BodyReturn = BodyReturn.replace("{{FootNote}}", footnoteReference);

                    FootNotesString += '<div style=\'mso-element:footnote\' id=ftn' + (i + 1) + '>'+
                    '<p class=MsoFootnoteText>'+
                    '<a style=\'mso-footnote-id:ftn' + (i + 1) + '\' href="#_ftnref' + (i + 1) + '"'+
                    ' name="_ftn' + (i + 1) + '" title=""><span class=MsoFootnoteReference><span style=\'mso-special-character:'+
                    'footnote\'><![if !supportFootnotes]><span class=MsoFootnoteReference><span'+
                    ' style=\'font-size:10.0pt;font-family:"Times New Roman",serif;mso-fareast-font-family:'+
                    '"Times New Roman";mso-fareast-theme-font:minor-fareast;mso-ansi-language:EN-US;'+
                    'mso-fareast-language:EN-US;mso-bidi-language:AR-SA\'>[' + (i + 1) + ']</span></span><![endif]></span></span></a>'+
                    FootNotesArray[i] +'</p></div>';
                }

                FootNotesString += '</div>';

                return {Body: BodyReturn, Footnotes: FootNotesString}
            }
            else{
                return {Body: Body, Footnotes: ''}
            }
        };
    return {
        GenerateWordFile: GenerateWordFile,
        setParameters: setParameters,
        imageBodyParser: imageBodyParser,
        image_to_base64: image_to_base64,
        generateTemplateHeadFooter:generateTemplateHeadFooter,
        calculateAspectRatioFit:calculateAspectRatioFit,
        generatePagingTemplate: generatePagingTemplate,
        generateTableOfContents: generateTableOfContents,
        addFootNotes: addFootNotes
    };
}();




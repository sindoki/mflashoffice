'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            console.log('Add-in is ready.');
        });
    });

})();

function attachFile() {
    var fileUrlJpg = 'https://amwine.ru/upload/blog/30-10-2019/5.jpg';
    var fileUrlDoc = 'https://win2.msoftgroup.ru/mflash/Dispatcher.php?C=1000&file_id=a9228562-ae77-11ea-afa9-525400868cf2&box=&user_id=2476&flpu=';
    var fileName = 'Test File.docx';

    console.log("Attach");

    Office.context.mailbox.item.addFileAttachmentAsync(fileUrlDoc, fileName, {}, (result) => {
        console.log(result)
    });
}
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
    var fileUrl = 'https://amwine.ru/upload/blog/30-10-2019/5.jpg';
    //var fileUrl = 'https://1drv.ms/w/s!AlK0_ugldxQ9gQg9Wya9CYoresIW?e=VjODug';
    //var fileUrl = 'https://1drv.ms/u/s!AlK0_ugldxQ9hxKlvp2iIlPnoewV';
    var fileName = 'File.jpg';

    console.log("Attach");

    Office.context.mailbox.item.addFileAttachmentAsync(fileUrl, fileName, {}, (result) => {
        console.log(result)
    });
}


<html>
    <head>
        <script type="text/javascript">
            function callback1(barcode, barcodeTypeId, barcodeTypeString) {
                var barcodeText = "BARCODE:" + barcode + " ID:" + barcodeTypeId + " STR:" + barcodeTypeString;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
            }

            function callback2( name, number, expMonth, expYear, track1, track2, track3 ) {
                var swipeText = "NAME:" + name + " NUMBER:" + number + " EXPMONTH:" + expMonth + " EXPYEAR:" + expYear + " TRACK1:" + track1 + " TRACK2:" + track2 + " TRACK3:" + track3;

                document.getElementById('swipe').innerHTML = swipeText;
                console.log(swipeText);
            }
            
            function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
                var barcodeText = "BARCODE:" + barcode + " ID:" + barcodeTypeId + " STR:" + barcodeTypeString;

                document.getElementById('barcode').innerHTML = barcodeText;
                console.log(barcodeText);
            }

            function adaptiscanSwipeFinished( name, number, expMonth, expYear, track1, track2, track3 ) {
                var swipeText = "NAME:" + name + " NUMBER:" + number + " EXPMONTH:" + expMonth + " EXPYEAR:" + expYear + " TRACK1:" + track1 + " TRACK2:" + track2 + " TRACK3:" + track3;

                document.getElementById('swipe').innerHTML = swipeText;
                console.log(swipeText);
            }

        </script>
    </head>
    <body>
        <h3>Adaptiscan Test Page</h3>
        <p>Scan a barcode or swipe a magnetic card.</p>
        <p>You can change the default home page by pressing the blue settings button in the top left corner.</p>
        <form>
            <span style="font-size:x-large;">BAR CODE</span>
            <div id="barcode" style="font-size:large;"></div>

            <span style="font-size:x-large;">SWIPE</span>
            <div id="swipe" style="font-size:large;"></div>

            <span style="font-size:x-large;">textarea1:</span>
            <textarea id="textarea1"></textarea>
            
            <br />
            <span style="font-size:x-large;">text1:</span>
            <input type="text" id="text1" />

            <br />
            <span style="font-size:x-large;">text2:</span>
            <input type="text" id="text2" />

            <br />
            <span style="font-size:x-large;">text3:</span>
            <input type="text" id="text3" />

            <br />
            <span style="font-size:x-large;">text4:</span>
            <input type="text" id="text4" />

            <br />
            <span style="font-size:x-large;">textarea2:</span>
            <textarea id="textarea2"></textarea>
        </form>
        
        <pre style="border: 1px solid #014D68; overflow: auto;">

        // The following javascript callback functions are implemented on this page.

        function callback1(barcode, barcodeTypeId, barcodeTypeString) {
            var barcodeText = "BARCODE:" + barcode + " ID:" + barcodeTypeId + " STR:" + barcodeTypeString;

            document.getElementById('barcode').innerHTML = barcodeText;
            console.log(barcodeText);
        }

        function callback2( name, number, expMonth, expYear, track1, track2, track3 ) {
            var swipeText = "NAME:" + name + " NUMBER:" + number + " EXPMONTH:" + expMonth + " EXPYEAR:" + expYear + " TRACK1:" + track1 + " TRACK2:" + track2 + " TRACK3:" + track3;

            document.getElementById('swipe').innerHTML = swipeText;
            console.log(swipeText);
        }

        function adaptiscanBarcodeFinished(barcode, barcodeTypeId, barcodeTypeString) {
            var barcodeText = "BARCODE:" + barcode + " ID:" + barcodeTypeId + " STR:" + barcodeTypeString;

            document.getElementById('barcode').innerHTML = barcodeText;
            console.log(barcodeText);
        }

        function adaptiscanSwipeFinished( name, number, expMonth, expYear, track1, track2, track3 ) {
            var swipeText = "NAME:" + name + " NUMBER:" + number + " EXPMONTH:" + expMonth + " EXPYEAR:" + expYear + " TRACK1:" + track1 + " TRACK2:" + track2 + " TRACK3:" + track3;

            document.getElementById('swipe').innerHTML = swipeText;
            console.log(swipeText);
        }

        </pre>
    </body>
</html>

<script type="text/javascript">

function getFirstTextInput() {
    var forms = document.forms || [];
    for(var i = 0; i < forms.length; i++){
        for(var j = 0; j < forms[i].length; j++){
            if(!forms[i][j].readonly != undefined
                    && (forms[i][j].type == "text" || forms[i][j].type == "textarea")
                    && forms[i][j].disabled != true && forms[i][j].style.display != 'none') {
				var id = forms[i][j].id;
                return forms[i][j].id;
            }
        }
    }
}

//getFirstTextInput();

</script>
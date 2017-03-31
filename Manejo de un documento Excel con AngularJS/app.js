/**
 * Created by Raúl Noa on 23/11/2016.
 */
(function(){
    'use strict';

    var app = angular.module('AngularExcel', ['ng', 'angular-js-xlsx', 'ngMdIcons', 'contenteditable'])

        
    .controller('AppCtrl', function ($scope) {
        $scope.name = "Select file";
        var name;
        var first_sheet_name;
        var jsonData;
        var newWorkBook;
        $scope.saveButton = true;
        var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];

        

        /**
         * Leer documento Excel
         */

        $scope.read = function (workbook) {

            $scope.saveButton = false;

            newWorkBook = workbook;

            first_sheet_name = newWorkBook.SheetNames[0];

            $scope.$apply(function(){

                jsonData = XLSX.utils.sheet_to_json(newWorkBook.Sheets[first_sheet_name]);
                jsonData[0]["rowIndex"] = "0";
                $scope.Elements = [jsonData[0]];

                for (var i = 1; i<jsonData.length;i++)
                {
                    $scope.Elements.push(
                        jsonData[i]
                    );
                    jsonData[i]["rowIndex"] = i;
                }
            });

        };

        /**
         * Salvar documento excel con los cambios realizados
         */
        $scope.saveData = function (){
            //  bookType can be 'xlsx' or 'xlsm' or 'xlsb'
            var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };

            var wbout = XLSX.write(newWorkBook,wopts);

            function s2ab(s) {
                var buf = new ArrayBuffer(s.length);
                var view = new Uint8Array(buf);
                for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                return buf;
            }

            /**
             * Capturando la fecha y la hora a la
             * que se realiza la salva
             */

            var d = new Date();

            var datestring = d.getDate()  + "-" + (d.getMonth()+1) + "-" + d.getFullYear() + " " +
                d.getHours() + "-" + d.getMinutes();

            /* the saveAs call downloads a file on the local machine */
            saveAs(new Blob([s2ab(wbout)],{type:""}), "New Excel file "+datestring+".xlsx");
        };

        $scope.changedCeld = function(val){
            var numberColumn = event.target.getAttribute('id');
            var numberArrow = event.target.parentElement.getAttribute('id');
            var worksheet = newWorkBook.Sheets[first_sheet_name];
            /**
             * Definimos la celda con la estructura LETRA + NÚMERO DE FILA
             * */
            var address_of_cell = alphabet[numberColumn]+(parseInt(numberArrow)+2);

            /**
             * Creamos la celda que contiene los datos modificados y
             * actualizamos los datos
             * */
            var newcell = {h: val, r:"<t>"+val+"</t>", t:"s", v:val, w:val};
            worksheet[address_of_cell] = newcell;

        };


        var drop = document.getElementById('drop');
        var workbook2;
        function handleDrop(e) {
            e.stopPropagation();
            e.preventDefault();

            var files = e.dataTransfer.files;
            var f = files[0];
            {
                var reader = new FileReader();
                name = f.name;

                reader.onload = function(evt) {
                    if(typeof console !== 'undefined')
                        var data = evt.target.result;

                    /* if binary string, read with type 'binary' */
                    try {
                        /*Mostrar el nombre del archivo cargado en el label*/
                        $scope.name = name;
                        $scope.saveButton = false;
                        $scope.$apply();

                        /*Leer el archivo cargado*/
                        workbook2 = XLS.read(data, {type: 'binary'});
                        $scope.read(workbook2);


                    } catch(e) {
                        console.log("error");
                        alert("This is not a excel file");
                    }
                };
                reader.readAsBinaryString(f);
            }
        }

        function handleDragover(e) {
            e.stopPropagation();
            e.preventDefault();
            e.dataTransfer.dropEffect = 'copy';
        }

        if(drop.addEventListener) {
            drop.addEventListener('dragenter', handleDragover, false);
            drop.addEventListener('dragover', handleDragover, false);
            drop.addEventListener('drop', handleDrop, false);
        }

    });

    /**
    * Para mostrar el nombre del archivo excel
    * cargado mediante la opción de búsqueda en el equipo
    */
    app.directive('fdInput', [function () {
        return {
            link: function ($scope, element, attrs) {
                element.on('change', function  (evt) {
                    var files = evt.target.files;

                    $scope.name = files[0].name;

                });
            }
        };
    }]);
    

})();
/**
 * Created by Raúl Noa on 23/11/2016.
 */
(function(){
    'use strict';

    var app = angular.module('AngularExcel', ['ng', 'angular-js-xlsx', 'ngMdIcons', 'contenteditable'])

        
    .controller('AppCtrl', function ($q, $scope) {
        $scope.name = "Select file";
        var name;
        var workbookNew;
        var first_sheet_name;
        var jsonData;
        var totalProducts;
        var newWorkBook;
        $scope.saveButton = true;
        

        /**
         * Leer documento Excel
         */

        $scope.read = function (workbook) {

            $scope.saveButton = false;

            newWorkBook = workbook;

            first_sheet_name = workbook.SheetNames[0];
            jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[first_sheet_name]);

            $scope.Elements = [jsonData[0]];
            $scope.$apply();

            for (var i = 1; i<jsonData.length;i++)
            {
                $scope.Elements.push(
                    jsonData[i]
                );
                $scope.$apply();
            }

        };

        /**
         * Salvar documento excel con los cambios realizados
         */
        $scope.saveData = function (){
          alert("This functionality do not work yet");
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
/**
 * Created by Raúl Noa on 23/11/2016.
 */
(function(){
    'use strict';

    var app = angular.module('AngularExcel', ['ng', 'angular-js-xlsx', 'ngMdIcons'])

        
    .controller('AppCtrl', function ($q, $scope) {
        $scope.name = "Select file";
        var name;
        var workbookNew;
        var first_sheet_name;
        var jsonData;
        var totalProducts;
        var workBookAutofill;
        

        /**
         * Read excel file
         */

        $scope.read = function (workbook) {

            first_sheet_name = workbook.SheetNames[0];
            jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[first_sheet_name]);

            $scope.Elements = [jsonData[0]];
            $scope.$apply();

            for (var i = 1; i<jsonData.length;i++)
            {
                for(var name in jsonData[i]) {
                 //   console.log(name);
                    var value = jsonData[i][name];
                   // console.log(value);
                }
                $scope.Elements.push(
                    jsonData[i]
                );
                $scope.$apply();

            }



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
                        workbook2 = XLS.read(data, {type: 'binary'});
                        $scope.read(workbook2);

                        $scope.name = name;

                    } catch(e) {
                        console.log("error");
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
    app.directive('fdInput', [function () {
        return {
            link: function ($scope, element, attrs) {
                element.on('change', function  (evt) {
                    var files = evt.target.files;
                  //  console.log(files[0].name);

                    $scope.name = files[0].name;

                });
            }
        };
    }]);
    

})();
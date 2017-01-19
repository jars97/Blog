/**
 * Created by Raúl Noa on 23/11/2016.
 */
(function(){
    'use strict';

    var app = angular.module('AngularExcel', ['ng', 'angular-js-xlsx'])

        
    .controller('AppCtrl', function ($q, $scope) {
        $scope.name = "Select file";
        var name;
        var workbookNew;
        var first_sheet_name;
        var jsonData;
        var totalProducts;
        var workBookAutofill;



        /*$scope.Elements = [

        ];*/


       /* $scope.gridOptions = {
            enableSorting: true,
            columnDefs: [
                { field: 'name' },
                { field: 'age' },
                { field: 'location', enableSorting: false }
            ],
            onRegisterApi: function (gridApi) {
                $scope.grid1Api = gridApi;
            }
        };*/

        /* $scope.gridOptions = {
            enableSorting: true,
            columnDefs: [],
            onRegisterApi: function (gridApi) {
                $scope.grid1Api = gridApi;
            }
        };

        $scope.gridOptions.data = [];*/


        /*$scope.gridOptions = {
            enableSorting: true,
            columnDefs: [
            ],
            onRegisterApi: function (gridApi) {
                $scope.grid1Api = gridApi;
            }
        };*/



      /*  $scope.myData = [{name: "Moroni", age: 50},
            {name: "Tiancum", age: 43},
            {name: "Jacob", age: 27},
            {name: "Nephi", age: 29},
            {name: "Enos", age: 34}];
        $scope.gridOptions = { data : 'myData' };*/

        /*$scope.users = [
            { name: "", age: '' , location: '' },

        ];*/

      /*  $scope.users = [
            { name: "Madhav Sai", age: 10, location: 'Nagpur' },
            { name: "Suresh Dasari", age: 30, location: 'Chennai' },
            { name: "Rohini Alavala", age: 29, location: 'Chennai' },
            { name: "Praveen Kumar", age: 25, location: 'Bangalore' },
            { name: "Sateesh Chandra", age: 27, location: 'Vizag' }
        ];
        $scope.gridOptions.data = $scope.users;*/

        /**
         * Read excel file
         */

        $scope.read = function (workbook) {



            first_sheet_name = workbook.SheetNames[0];
            jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[first_sheet_name]);
            console.log(jsonData);
            console.log(jsonData[0]);
            //console.log(jsonData[0][1]);

            /*var size = Object.keys(jsonData[0]).length;
            console.log("dddd "+size);

            $scope.Elements.push(
                jsonData
            );
            $scope.$apply();*/

            /*$scope.Elements = [
                { value: '1', url: '' },
                { value: '2', url: '' },
                { value: '3', url: '' },
                { value: '4', url: '' },
                { value: '5', url: '' },
                { value: '6', url: '' },
                { value: '7', url: '' },
                { value: '8', url: '' },
                { value: '9', url: '' }
            ];*/

            /*$scope.Elements.push(
                jsonData[0]
            );*/

            /*$scope.Elements ={
                "Name" : "Alfreds Futterkiste",
                "Country" : "Germany",
                "City" : "Berlin"
            };*/

            console.log(jsonData.length);
            /*$scope.Elements = [jsonData[0]];
            $scope.$apply();*/

            for (var i = 0; i<jsonData.length;i++)
            {
                for(var name in jsonData[i]) {
                    console.log(name);
                    var value = jsonData[i][name];
                    console.log(value);
                }
                /*$scope.Elements.push(
                    jsonData[i]
                );
                $scope.$apply();*/

            }


            /*$scope.Elements = jsonData[0];

            $scope.$apply();*/

            /*$scope.gridOptions = {
                enableSorting: true,
                columnDefs: [
                    { field: 'name' },
                    { field: 'age' },
                    { field: 'location', enableSorting: false }
                ],
                onRegisterApi: function (gridApi) {
                    $scope.grid1Api = gridApi;
                }
            };*/

         /*   $scope.users = [
                { name: "hgjjgh", age: 10, location: 'gjghj' },
                { name: "Suresh ghjhg", age: 30, location: 'Chenhjhgjnai' },
                { name: "Rohini Alavala", age: 29, location: 'Cheghjghjnnai' },
                { name: "Praveen Kumar", age: 25, location: 'Bangalore' },
                { name: "Sateesh Chandra", age: 27, location: 'Vizag' }
            ];

            $scope.gridOptions.data = $scope.users;*/

          //  console.log("cghchgf");
         //  $scope.gridApi.core.refresh();

            /* DO SOMETHING WITH workbook HERE */
          /*  $scope.noexist = 0;
            $scope.disabled = false;
            workbookNew = workbook;

            /!*Changing sheet columns name*!/
            first_sheet_name = workbookNew.SheetNames[0];
            var address_of_cell_Google = 'I1';
            var address_of_cell_Description = 'J1';
            var worksheet = workbookNew.Sheets[first_sheet_name];

            var newcellGoogle = {h: "Image (from Google source)", r:"<t>Image (from Google source)</t>", t:"s", v:"Image (from Google source)", w:"Image (from Google source)"};
            var newcellDescription = {h: "English Product Description (Alzashop.com)", r:"<t>English Product Description (Alzashop.com)</t>", t:"s", v:"English Product Description (Alzashop.com)", w:"English Product Description (Alzashop.com)"};
            worksheet[address_of_cell_Google] = newcellGoogle;
            worksheet[address_of_cell_Description] = newcellDescription;

            for (var sheetName in workbookNew.Sheets) {
                jsonData = XLSX.utils.sheet_to_json(workbookNew.Sheets[sheetName]);
            }



            /!**
             * Autofill excel file
             *!/
            console.log(jsonData);
            console.log(jsonData[0]);

            var lolo = jsonData[0];
          //  console.log(lolo['Coal']);

            angular.forEach(lolo, function(value, key) {
                console.log(key);
            });

        //   $scope.myData = jsonData;
           /!* $scope.myData = [{lolo: "Noa", age: 50},
                {lolo: "Noa", age: 43},
                {lolo: "Jacob", age: 27},
                {lolo: "Nephi", age: 29},
                {lolo: "Enos", age: 34}];
            $scope.gridOptions = { data : 'myData' };*!/

            $scope.users = [
                { name: "Madhav Sai", age: 10, location: 'Nagpur' },
                { name: "Suresh Dasari", age: 30, location: 'Chennai' },
                { name: "Rohini Alavala", age: 29, location: 'Chennai' },
                { name: "Praveen Kumar", age: 25, location: 'Bangalore' },
                { name: "Sateesh Chandra", age: 27, location: 'Vizag' }
            ];*/

          //  addRestProducts(jsonData);


            // addRestProducts(jsonData);
            //  workbookNew.SheetNames[0] ="Alzashop";

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
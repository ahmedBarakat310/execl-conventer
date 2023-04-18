let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0]
 let rowInput= document.getElementById('row')
 let columnInput= document.getElementById('col')
tableExists = false


const generateTable = () => {

    let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value)
    if(rowsNumber<1 || rowInput.value=='' ){
      Swal.fire('please enter the number of table rows')
    }else{
        if(columnsNumber<1 || columnInput.value==''){
            Swal.fire('please enter the number of table columns')
        }
        table.innerHTML = ""
        for(let i=0; i<rowsNumber; i++){
            var tableRow = ""
            for(let j=0; j<columnsNumber; j++){
                tableRow += `<td contenteditable></td>`
            }
            table.innerHTML += tableRow
        }
        if(rowsNumber>0 && columnsNumber>0){
            tableExists = true
        }
    }
    
    }
    let name1='defult';
    window.localStorage.setItem('nameFile',name1+".")
 let nameOfFileInput=document.getElementById('nameOfFile');
    nameOfFileInput.addEventListener('input',function(e){
        name1=e.target.value
        
        window.localStorage.setItem('nameFile',name1+".")
      
})


const ExportToExcel = (type, fn, dl) => {
    if(!tableExists){
        Swal.fire('please generate the table first')
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || (window.localStorage.getItem('nameFile') + (type || 'xlsx')))
}
import React, { Component } from "react";
import * as XLSX from 'xlsx';

export default class Excel extends Component {

    state = {
        woorksheets: [],
        filas: [],
        propiedades: [],
        status: false
    }

    selectHoja = React.createRef();

    leerFilas = (index) => {
        var hoja = this.state.woorksheets[index];
        var filas = XLSX.utils.sheet_to_row_object_array(hoja.data);
        this.state.filas = [];
        this.state.filas = filas;
    }

    leerPropiedades = (index) => {
        var hoja = this.state.woorksheets[index];
        this.state.propiedades = [];
        for (let key in hoja.data) {
            let regEx = new RegExp("^\(\\w\)\(1\){1}$");
            if (regEx.test(key) == true) {
                this.state.propiedades.push(hoja.data[key].v);
            }
        }
    }

    cambiarHoja = () => {
        this.leerPropiedades(this.selectHoja.current.value);
        this.leerFilas(this.selectHoja.current.value);
        this.setState({
            filas: this.state.filas,
            propiedades: this.state.propiedades
        })
    }

    leerExcel = (e) => {
        e.preventDefault();
        const formData = new FormData(e.currentTarget);
        var excel = formData.get("excel");
        var listWorksheet = [];

        var reader = new FileReader()
        reader.readAsArrayBuffer(excel)
        reader.onloadend = (e) => {
            var data = new Uint8Array(e.target.result)
            var excelRead = XLSX.read(data, {type: 'array'})
            excelRead.SheetNames.forEach(function(sheetName, index) {
                listWorksheet.push({
                    data: excelRead.Sheets[sheetName], 
                    name: sheetName, 
                    index: index
                })
            })
            
            this.state.woorksheets = listWorksheet;
            this.setState({
                woorksheets: this.state.woorksheets
            })

            this.leerPropiedades(0);
            this.leerFilas(0);
            this.setState({
                filas: this.state.filas,
                propiedades: this.state.propiedades,
                status: true
            })
        }        
    }

    render() {
        return (
        <div className="container">
            <h1 className="my-5">Leer excel</h1>
            <form onSubmit={this.leerExcel}>
                <label className="form-label">Selecciona un archivo excel: </label>
                <input type={"file"} className="form-control" 
                accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
                name="excel"/>
                <button className="btn btn-primary mt-3">Convertir</button>
            </form>
            <hr/>
            {
                this.state.status &&
                <>
                    <form>
                        <label className="form-label">Hojas </label>
                        <select className='form-select' ref={this.selectHoja} onChange={this.cambiarHoja}>
                        {
                            this.state.woorksheets.map((hoja, index) => {
                                return (<option key={index} value={index}>{hoja.name}</option>)
                            })
                        }
                        </select>
                    </form>
                    <table className="table mt-3">
                        <thead>
                            <tr>
                            {
                                this.state.propiedades.map((propiedad, index) => {                                    
                                    return (                                      
                                            <th key={index}>
                                                {propiedad}
                                            </th>                                        
                                    )                                    
                                })
                            }
                            </tr>
                        </thead>
                        <tbody>
                            {
                                this.state.filas.map((fila, index1) => {
                                    return (
                                    <tr key={index1}>
                                        {
                                            this.state.propiedades.map((propiedad, index2) => {
                                                return <td>{fila[propiedad]}</td>
                                            })
                                        }
                                    </tr>
                                    )
                                })

                            }
                        </tbody>
                    </table>                
                </>
            }
        </div>
        );
    }
}

import React, { useState } from 'react';
import axios from 'axios';
import { useTable, useFilters, useSortBy } from 'react-table';
import * as XLSX from 'xlsx';

const Dashboard = () => {
    const [data, setData] = useState([]);
    const [fileImport, setFileImport] = useState(null);
    const [fileExtract, setFileExtract] = useState(null);
    const [facturationPercentage, setFacturationPercentage] = useState(4);
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState(null);

    const handleFileChange = (e, type) => {
        const file = e.target.files[0];
        if (file) {
            if (type === 'import') setFileImport(file);
            if (type === 'extract') setFileExtract(file);
        }
    };

    const handleUpload = async () => {
        if (!fileImport || !fileExtract) {
            setError("Veuillez sélectionner les deux fichiers");
            return;
        }

        setIsLoading(true);
        setError(null);

        try {
            const formData = new FormData();
            formData.append('import', fileImport);
            formData.append('extract', fileExtract);
            
            const response = await axios.post('http://localhost:5000/upload', formData, {
                headers: { 'Content-Type': 'multipart/form-data' },
            });
            
            const updatedData = response.data.processedData.map(row => ({
                ...row,
                "Cu's à facturer": (row["Total_Cu+"] / (facturationPercentage / 100)).toFixed(2),
                "Cu's à retirer": (row["Total_Cu+"] / (facturationPercentage / 100) - row["Total général"]).toFixed(2)
            }));

            setData(updatedData);
        } catch (err) {
            setError("Erreur lors du traitement des fichiers: " + err.message);
        } finally {
            setIsLoading(false);
        }
    };

    const exportToExcel = () => {
        if (data.length === 0) {
            setError("Aucune donnée à exporter");
            return;
        }

        const exportData = data.map(row => ({
            "Agent": row.agent,
            "Total général": Number(row["Total général"]),
            "Don avec montant": Number(row["don avec montant"]),
            "Don en ligne": Number(row["don en ligne"]),
            "Total Cu+": Number(row["Total_Cu+"]),
            "Tx Accord don": row["Tx Accord_don"],
            "PA": Number(row["pa"]),
            "Pa en ligne": Number(row["pa en ligne"]),
            "Tx Accord Pal": row["Tx Accord_Pal"],
            "Indécis Don": Number(row["indecis Don"]),
            "Refus argumenté": Number(row["refus argumente"]),
            "Durée production": Number(row["Durée production"]),
            "Durée présence": Number(row["Durée présence"]),
            "Cu's/h": Number(row["Cu's/h"]),
            "Nbr/J Travailler": Number(row["Nbr/J Travailler"]),
            "Cu's à facturer": Number(row["Cu's à facturer"]),
            "Cu's à retirer": Number(row["Cu's à retirer"])
        }));

        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        
        // Configuration du format des colonnes
        worksheet['!cols'] = [
            { wch: 20 }, // Agent
            { wch: 12 }, // Total général
            { wch: 15 }, // Don avec montant
            { wch: 12 }, // Don en ligne
            { wch: 12 }, // Total Cu+
            { wch: 12 }, // Tx Accord don
            { wch: 10 }, // PA
            { wch: 12 }, // Pa en ligne
            { wch: 12 }, // Tx Accord Pal
            { wch: 12 }, // Indécis Don
            { wch: 15 }, // Refus argumenté
            { wch: 15 }, // Durée production
            { wch: 15 }, // Durée présence
            { wch: 10 }, // Cu's/h
            { wch: 15 }, // Nbr/J Travailler
            { wch: 15 }, // Cu's à facturer
            { wch: 15 }  // Cu's à retirer
        ];

        XLSX.utils.book_append_sheet(workbook, worksheet, "Résultats");
        XLSX.writeFile(workbook, `tableau_resultats_${new Date().toISOString().split('T')[0]}.xlsx`);
    };

    const columns = React.useMemo(() => [
        { Header: 'Agent', accessor: 'agent' },
        { Header: 'Total général', accessor: 'Total général' },
        { Header: 'Don avec montant', accessor: 'don avec montant' },
        { Header: 'Don en ligne', accessor: 'don en ligne' },
        { Header: 'Total Cu+', accessor: 'Total_Cu+' },
        { Header: 'Tx Accord don', accessor: 'Tx Accord_don' },
        { Header: 'PA', accessor: 'pa' },
        { Header: 'Pa en ligne', accessor: 'pa en ligne' },
        { Header: 'Tx Accord Pal', accessor: 'Tx Accord_Pal' },
        { Header: 'Indécis Don', accessor: 'indecis Don' },
        { Header: 'Refus argumenté', accessor: 'refus argumente' },
        { Header: 'Durée production', accessor: 'Durée production' },
        { Header: 'Durée présence', accessor: 'Durée présence' },
        { Header: "Cu's/h", accessor: "Cu's/h" },
        { Header: 'Nbr/J Travailler', accessor: 'Nbr/J Travailler' },
        { Header: "Cu's à facturer", accessor: "Cu's à facturer" },
        { Header: "Cu's à retirer", accessor: "Cu's à retirer" }
    ], []);

    const {
        getTableProps,
        getTableBodyProps,
        headerGroups,
        rows,
        prepareRow
    } = useTable(
        {
            columns,
            data,
            initialState: {
                sortBy: [{ id: 'agent', desc: false }]
            }
        },
        useFilters,
        useSortBy
    );

    return (
        <div className="p-5">
            <h1 className="text-2xl font-bold mb-4">Tableau des Résultats</h1>
            
            <div className="mb-4 space-y-4">
                <div className="flex flex-wrap gap-4 items-center">
                    <div className="flex-1 min-w-[200px]">
                        <label className="block text-sm font-medium mb-1">Fichier Import</label>
                        <input 
                            type="file" 
                            onChange={(e) => handleFileChange(e, 'import')}
                            accept=".xlsx,.xls"
                            className="w-full p-2 border rounded"
                        />
                    </div>
                    <div className="flex-1 min-w-[200px]">
                        <label className="block text-sm font-medium mb-1">Fichier Extract</label>
                        <input 
                            type="file" 
                            onChange={(e) => handleFileChange(e, 'extract')}
                            accept=".xlsx,.xls"
                            className="w-full p-2 border rounded"
                        />
                    </div>
                    <div className="min-w-[150px]">
                        <label className="block text-sm font-medium mb-1">Pourcentage</label>
                        <select 
                            onChange={(e) => setFacturationPercentage(Number(e.target.value))}
                            value={facturationPercentage}
                            className="w-full p-2 border rounded"
                        >
                            {[4, 5, 6, 7, 8, 9, 10, 11, 12].map((percent) => (
                                <option key={percent} value={percent}>{percent}%</option>
                            ))}
                        </select>
                    </div>
                </div>

                <div className="flex gap-4">
                    <button 
                        onClick={handleUpload}
                        disabled={isLoading || !fileImport || !fileExtract}
                        className={`px-4 py-2 rounded ${
                            isLoading || !fileImport || !fileExtract
                                ? 'bg-gray-400'
                                : 'bg-blue-500 hover:bg-blue-600'
                        } text-white transition-colors`}
                    >
                        {isLoading ? 'Traitement...' : 'Importer'}
                    </button>
                    <button 
                        onClick={exportToExcel}
                        disabled={data.length === 0}
                        className={`px-4 py-2 rounded ${
                            data.length === 0
                                ? 'bg-gray-400'
                                : 'bg-green-500 hover:bg-green-600'
                        } text-white transition-colors`}
                    >
                        Télécharger Excel
                    </button>
                </div>

                {error && (
                    <div className="p-3 bg-red-100 border border-red-400 text-red-700 rounded">
                        {error}
                    </div>
                )}
            </div>

            <div className="overflow-x-auto shadow-md rounded-lg">
                <table {...getTableProps()} className="min-w-full divide-y divide-gray-200">
                    <thead className="bg-gray-50">
                        {headerGroups.map((headerGroup, index) => (
                            <tr {...headerGroup.getHeaderGroupProps()} key={index}>
                                {headerGroup.headers.map((column, colIndex) => (
                                    <th
                                        {...column.getHeaderProps(column.getSortByToggleProps())}
                                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100"
                                        key={colIndex}
                                    >
                                        <div className="flex items-center">
                                            {column.render('Header')}
                                            <span className="ml-2">
                                                {column.isSorted
                                                    ? column.isSortedDesc
                                                        ? '▼'
                                                        : '▲'
                                                    : ''}
                                            </span>
                                        </div>
                                    </th>
                                ))}
                            </tr>
                        ))}
                    </thead>
                    <tbody {...getTableBodyProps()} className="bg-white divide-y divide-gray-200">
                        {rows.map((row, rowIndex) => {
                            prepareRow(row);
                            return (
                                <tr {...row.getRowProps()} className="hover:bg-gray-50" key={rowIndex}>
                                    {row.cells.map((cell, cellIndex) => (
                                        <td
                                            {...cell.getCellProps()}
                                            className="px-6 py-4 whitespace-nowrap text-sm text-gray-500"
                                            key={cellIndex}
                                        >
                                            {cell.render('Cell')}
                                        </td>
                                    ))}
                                </tr>
                            );
                        })}
                    </tbody>
                </table>
            </div>

            {data.length === 0 && !isLoading && (
                <div className="text-center py-4 text-gray-500">
                    Aucune donnée disponible
                </div>
            )}
        </div>
    );
};

export default Dashboard;
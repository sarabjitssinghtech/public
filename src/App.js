import React, { useEffect } from 'react';
import { useTable } from 'react-table';
import * as Excel from 'exceljs';

function App() {
  const data = React.useMemo(
    () => [
      {
        col1: 'Hello',
        col2: 'World',
      },
      {
        col1: 'react-table',
        col2: 'rocks',
      },
      {
        col1: 'whatever',
        col2: 'you want',
      },
    ],
    []
  );

  const columns = React.useMemo(
    () => [
      {
        Header: 'Column 1',
        accessor: 'col1', // accessor is the "key" in the data
      },
      {
        Header: 'Column 2',
        accessor: 'col2',
      },
    ],
    []
  );

  const { getTableProps, getTableBodyProps, headerGroups, rows, prepareRow } =
    useTable({ columns, data });

  useEffect(() => {
    async function loadData() {
      var workbook = new Excel.Workbook();
      workbook.creator = 'Web';
      workbook.lastModifiedBy = 'Web';
      workbook.created = new Date();
      workbook.modified = new Date();
      workbook.addWorksheet('SheetTest', {
        views: [
          {
            state: 'frozen',
            ySplit: 3,
            xSplit: 2,
            activeCell: 'A1',
            showGridLines: false,
          },
        ],
      });
      var sheet = workbook.getWorksheet(1);
      var head1 = ['Exported Data'];
      sheet.addRow(head1);
      sheet.addRow('');
      sheet.getRow(3).values = [
        'Column1',
        'Column2',
        'Column3',
        'Column4',
        'Column5',
      ];
      sheet.columns = [
        { key: 'col1' },
        { key: 'col2' },
        { key: 'col3' },
        { key: 'col4' },
        { key: 'col5' },
      ];
      const dataObj = [
        { col1: 'a1', col2: 'b1', col3: 'c1', col4: 'd1', col5: 'e1' },
        { col1: 'a2', col2: 'b2', col3: 'c2', col4: 'd2', col5: 'e2' },
        { col1: 'a3', col2: 'b3', col3: 'c3', col4: 'd3', col5: 'e3' },
        { col1: 'a4', col2: 'b4', col3: 'c4', col4: 'd4', col5: 'e4' },
        { col1: 'a5', col2: 'b5', col3: 'c5', col4: 'd5', col5: 'e5' },
      ];
      sheet.addRows(dataObj);
      workbook.xlsx.writeBuffer().then((data) => {
        var blob = new Blob([data], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
        });
        var url = window.URL.createObjectURL(blob);
        var a = document.createElement('a');
        document.body.appendChild(a);
        a.href = url;
        a.download = 'TestExcelExport.xlsx';
        a.click();
        //adding some delay in removing the dynamically created link solves the problem in FireFox
        //setTimeout(function() {window.URL.revokeObjectURL(url);},0);
        //FileSaver.saveAs(blob, this.excelFileName, true);
      });
    }
    loadData();
  }, []);

  return (
    <table {...getTableProps()} style={{ border: 'solid 1px blue' }}>
      <thead>
        {headerGroups.map((headerGroup) => (
          <tr {...headerGroup.getHeaderGroupProps()}>
            {headerGroup.headers.map((column) => (
              <th
                {...column.getHeaderProps()}
                style={{
                  borderBottom: 'solid 3px red',
                  background: 'aliceblue',
                  color: 'black',
                  fontWeight: 'bold',
                }}
              >
                {column.render('Header')}
              </th>
            ))}
          </tr>
        ))}
      </thead>
      <tbody {...getTableBodyProps()}>
        {rows.map((row) => {
          prepareRow(row);
          return (
            <tr {...row.getRowProps()}>
              {row.cells.map((cell) => {
                return (
                  <td
                    {...cell.getCellProps()}
                    style={{
                      padding: '10px',
                      border: 'solid 1px gray',
                      background: 'papayawhip',
                    }}
                  >
                    {cell.render('Cell')}
                  </td>
                );
              })}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

export default App;

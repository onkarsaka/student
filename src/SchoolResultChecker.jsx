import React, { useState, useMemo } from 'react';
import { useTable } from 'react-table';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import './SchoolResultChecker.css';
import schoollogo from "./schoollogo.jpg"

export default function SchoolResultChecker() {
    const [students, setStudents] = useState([]);
    const [newStudent, setNewStudent] = useState({ name: '' });
    const [subjects, setSubjects] = useState(['विषय १']);
    const [newSubjects, setNewSubjects] = useState('');

    const calculateGrade = (total) => {
        if (total >= 91) return 'अ १';
        if (total >= 81) return 'अ २';
        if (total >= 71) return 'ब १';
        if (total >= 61) return 'ब २';
        if (total >= 51) return 'क १';
        if (total >= 41) return 'क २';
        if (total >= 31) return 'ड';
        if (total >= 21) return 'इ १';
        return 'इ २';
    };

    const calculateGradeDescription = (grade) => {
        if (grade == 'अ १') return 'अप्रतिम'
        if (grade == 'अ २') return 'खुप चांगला'
        if (grade == 'ब १') return 'चांगला'
        if (grade == 'ब २') return 'बरा'
        if (grade == 'क १') return 'सर्वसाधारण'
        if (grade == 'क २') return 'ठीक'
        if (grade == 'ड') return 'असमानधारक'
        if (grade == 'इ १') return 'सुधारणा आवश्यक'
        return 'सुधारणा आवश्यक';
    }

    const calculatePassOrFail = (grade) => {
        if (grade == 'इ २') return 'नापास'
        return 'पास'
    }

    const calculateStats = (student) => {
        const subjectStats = subjects.reduce((acc, subject) => {
            const theory = Math.min(Number(student[`${subject}Theory`]) || 0, 100);
            const practical = Math.min(Number(student[`${subject}Practical`]) || 0, 100);
            const total = theory + practical;
            const grade = calculateGrade(total);
            acc[subject] = { theory, practical, total, grade };
            return acc;
        }, {});

        const overallTotal = Object.values(subjectStats).reduce((sum, { total }) => sum + total, 0);
        const overallAverage = overallTotal / subjects.length;
        const overallGrade = calculateGrade(overallAverage);
        const overallGradeDescription = calculateGradeDescription(overallGrade);
        const PassOrFail = calculatePassOrFail(overallGrade);

        return {
            ...student,
            ...subjectStats,
            overallTotal: overallTotal.toFixed(2),
            overallAverage: overallAverage.toFixed(2),
            overallGrade,
            overallGradeDescription,
            PassOrFail,
        };
    };

    const data = useMemo(() => students.map(calculateStats), [students, subjects]);

    const columns = useMemo(
        () => [
            { Header: 'विद्यार्थ्यांचे नाव', accessor: 'name' },
            ...subjects.map(subject => ({
                Header: subject,
                columns: [
                    {
                        Header: 'आकारीक मुल्य',
                        accessor: `${subject}Theory`,
                    },
                    {
                        Header: 'संकलित मुल्य',
                        accessor: `${subject}Practical`,
                    },
                    {
                        Header: 'एकूण',
                        accessor: `${subject}.total`,
                        Cell: ({ value }) => value?.toFixed(2) || '',
                    },
                    {
                        Header: 'श्रेणी',
                        accessor: `${subject}.grade`,
                    },
                ],
            })),
            { Header: 'एकूण गुण', accessor: 'overallTotal' },
            { Header: 'शेकडा प्रमाण', accessor: 'overallAverage' },
            { Header: 'श्रेणी', accessor: 'overallGrade' },
            { Header: 'श्रेणीवर्णन', accessor: 'overallGradeDescription' },
            { Header: 'शेरा', accessor: 'PassOrFail' },
            {
                Header: 'Actions',
                Cell: ({ row }) => (
                    <button onClick={() => deleteRow(row.original.id)} className="delete-row">
                        Delete
                    </button>
                ),
            },
        ],
        [subjects]
    );

    const {
        getTableProps,
        getTableBodyProps,
        headerGroups,
        rows,
        prepareRow,
    } = useTable({ columns, data });

    const handleInputChange = (e) => {
        setNewStudent({ ...newStudent, [e.target.name]: e.target.value });
    };

    const addStudent = () => {
        if (newStudent.name.trim()) {
            const studentData = {
                id: Date.now(),
                name: newStudent.name.trim(),
                ...subjects.reduce((acc, subject) => ({
                    ...acc,
                    [`${subject}Theory`]: '',
                    [`${subject}Practical`]: '',
                }), {}),
            };
            setStudents(prevStudents => [...prevStudents, studentData]);
            setNewStudent({ name: '' });
        }
    };

    const handleNewSubjectsChange = (e) => {
        setNewSubjects(e.target.value);
    };

    const addSubjects = () => {
        const subjectsToAdd = newSubjects.split(',').map(s => s.trim()).filter(s => s && !subjects.includes(s));
        setSubjects(prevSubjects => [...prevSubjects, ...subjectsToAdd]);
        setStudents(prevStudents => prevStudents.map(student => ({
            ...student,
            ...subjectsToAdd.reduce((acc, subject) => ({
                ...acc,
                [`${subject}Theory`]: '',
                [`${subject}Practical`]: '',
            }), {}),
        })));
        setNewSubjects('');
    };

    const removeSubjects = () => {
        const subjectsToRemove = newSubjects.split(',').map(s => s.trim());
        setSubjects(prevSubjects => prevSubjects.filter(subject => !subjectsToRemove.includes(subject)));
        setStudents(prevStudents => prevStudents.map(student => {
            const updatedStudent = { ...student };
            subjectsToRemove.forEach(subject => {
                delete updatedStudent[`${subject}Theory`];
                delete updatedStudent[`${subject}Practical`];
            });
            return updatedStudent;
        }));
        setNewSubjects('');
    };

    const handleCellEdit = (studentId, field, value) => {
        setStudents(prevStudents => prevStudents.map(student =>
            student.id === studentId ? { ...student, [field]: value } : student
        ));
    };

    const deleteRow = (id) => {
        setStudents(prevStudents => prevStudents.filter(student => student.id !== id));
    };

    const exportToExcel = () => {
        const wsData = [];
    
        const schoolName = [document.getElementById('school-name').innerText];
        const teacherName = [document.getElementById('teacher-name').innerText];
        const schoolYear = [document.getElementById('school-year').innerText];
        const schoolClass = [document.getElementById('school-class').innerText];
        const schoolSection = [document.getElementById('school-section').innerText];
        const schoolPaperHeader = [document.getElementById('school-paper-header').innerText];
        const schoolSemester = [document.getElementById('school-semester').innerText];
    
        wsData.push(schoolName);
        wsData.push(teacherName);
        wsData.push(schoolYear);
        wsData.push(schoolClass);
        wsData.push(schoolSection);
        wsData.push(schoolSemester);
        wsData.push(schoolPaperHeader);
        wsData.push('');

        const mergeRanges = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }, 
            { s: { r: 1, c: 0 }, e: { r: 1, c: 0 } },
            { s: { r: 2, c: 0 }, e: { r: 2, c: 0 } }, 
            { s: { r: 3, c: 0 }, e: { r: 3, c: 0 } }, 
            { s: { r: 4, c: 0 }, e: { r: 4, c: 0 } }, 
            { s: { r: 5, c: 0 }, e: { r: 5, c: 0 } }, 
            { s: { r: 6, c: 0 }, e: { r: 6, c: 0 } }, 
        ];
    
        const headerRow1 = ['विद्यार्थ्यांचे नाव'];
        subjects.forEach((subject, index) => {
            headerRow1.push(subject, '', '', ''); 
            const colStart = 1 + index * 4;
            const colEnd = colStart + 3;
            mergeRanges.push({ s: { r: 8, c: colStart }, e: { r: 8, c: colEnd } });
        });
    
        headerRow1.push('एकूण गुण', 'शेकडा प्रमाण', 'श्रेणी', 'श्रेणीवर्णन', 'शेरा');
        wsData.push(headerRow1);
    
        const headerRow2 = [''];
        subjects.forEach(() => {
            headerRow2.push('आकारीक मुल्य', 'संकलित मुल्य', 'एकूण', 'श्रेणी');
        });
        headerRow2.push(' ', ' ', ' ');
        wsData.push(headerRow2);
    
        data.forEach((student) => {
            const row = [student.name];
    
            subjects.forEach((subject) => {
                const theory = student[`${subject}Theory`] || '';
                const practical = student[`${subject}Practical`] || '';
                const total = Number(theory) + Number(practical) || 'NaN';
                const grade = calculateGrade(total) || 'NaN';
                row.push(theory, practical, total, grade);
            });
    
            row.push(
                student.overallTotal || '',
                (student.overallTotal / subjects.length).toFixed(2) || '',
                student.overallGrade || '',
                student.overallGradeDescription,
                student.PassOrFail,
            );
            wsData.push(row);
        });
    
        const ws = XLSX.utils.aoa_to_sheet(wsData);
    
        ws['!merges'] = mergeRanges;
    
        mergeRanges.forEach((range) => {
            for (let r = range.s.r; r <= range.e.r; r++) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cell = ws[XLSX.utils.encode_cell({ r, c })] || {};
                    cell.s = { alignment: { horizontal: 'center', vertical: 'center' } };
                    ws[XLSX.utils.encode_cell({ r, c })] = cell;
                }
            }
        });
    
        const maxColWidths = wsData[0].map((_, colIndex) =>
            Math.max(...wsData.map((row) => (row[colIndex] ? row[colIndex].toString().length : 0)))
        );
    
        ws['!cols'] = maxColWidths.map((width) => ({ width: width + 2 }));
    
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Results');
    
        XLSX.writeFile(wb, 'school_results.xlsx');
    };
    

    const exportToImage = () => {
        const table = document.getElementById('resultsTable');
        if (table) {
            html2canvas(table).then(canvas => {
                const link = document.createElement('a');
                link.download = 'school_results.png';
                link.href = canvas.toDataURL();
                link.click();
            });
        }
    };

    let isShiftAEnabled = false;

    document.addEventListener('keydown', (event) => {
        if (event.shiftKey && (event.key === 'a' || event.key === 'A')) {
            isShiftAEnabled = true;
        }

        if (event.key === 'Enter') {
            const activeElement = document.activeElement;
            if (activeElement.classList.contains('editable-header')) {
                event.preventDefault();
                activeElement.setAttribute('contenteditable', 'false');
            }
        }
    });

    document.addEventListener('keyup', (event) => {
        if (event.key === 'a' || event.key === 'A') {
            isShiftAEnabled = false;
        }
    });

    const headers = document.querySelectorAll('.editable-header');

    headers.forEach(header => {
        header.addEventListener('click', (event) => {
            if (isShiftAEnabled) {
                header.setAttribute('contenteditable', 'true');
                header.focus();
            }
        });

        header.addEventListener('blur', () => {
            header.setAttribute('contenteditable', 'false');
        });
    });

    return (
        <main>
            <section className='header-hero'>
                <img src={schoollogo} alt="" />
                <div>
                    <h1>कै.आ.ह.आब्बा प्राथमिक विद्यालय सोलापूर</h1>
                    <h3>94/41, जोडभावी पेठ सोलापूर</h3>
                </div>
            </section>

            <div className="school-result-checker" id="resultsTable">
                <div className='school-informations'>
                    <div className='school-informations-section1'>
                        <h3 id='school-name'>शाळेचे नांव : <span className="editable-header" contenteditable="false">कै.आ.ह.आब्बा प्राथमिक विद्यालय सोलापूर</span></h3>
                        <h3 id='teacher-name'>वर्ग शिक्षकाचे नांव : <span className="editable-header" contenteditable="false">शिक्षकांचे नाव</span></h3>
                    </div>
                    <div className='school-informations-section2'>
                        <h3 id='school-year'>सन : <span className="editable-header" contenteditable="false">२०२३-२४</span></h3>
                        <h3 id='school-class'>वर्ग : <span className="editable-header" contenteditable="false">१ ली</span></h3>
                        <h3 id='school-section'>तुकडी : <span className="editable-header" contenteditable="false">अ</span></h3>
                    </div>
                </div>
                <h1 id='school-paper-header' className="editable-header" contenteditable="false">सातत्यपूर्ण सर्वंकष मूल्यमापन</h1>
                <h3 id='school-semester' className="editable-header semester-header" contenteditable="false">प्रथम सत्र / द्वितीय सत्र</h3>
                <div className="input-section">
                    <input
                        type="text"
                        name="name"
                        placeholder="विद्यार्थ्यांचे नाव"
                        value={newStudent.name}
                        onChange={handleInputChange}
                    />
                    <button onClick={addStudent}>विद्यार्थी ॲड करा</button>
                    <input
                        type="text"
                        placeholder="विषय ॲड / रिमूव करा (comma-separated)"
                        value={newSubjects}
                        onChange={handleNewSubjectsChange}
                    />
                    <button onClick={addSubjects}>विषय ॲड करा</button>
                    <button onClick={removeSubjects}>विषय रिमूव करा</button>
                </div>
                <div className="table-container">
                    <table {...getTableProps()}>
                        <thead>
                            {headerGroups.map((headerGroup, i) => (
                                <tr {...headerGroup.getHeaderGroupProps()} key={i}>
                                    {headerGroup.headers.map((column, j) => (
                                        <th {...column.getHeaderProps()} key={j} colSpan={column.columns ? column.columns.length : 1}>
                                            {column.render('Header')}
                                        </th>
                                    ))}
                                </tr>
                            ))}
                        </thead>
                        <tbody {...getTableBodyProps()}>
                            {rows.map(row => {
                                prepareRow(row);
                                return (
                                    <tr {...row.getRowProps()}>
                                        {row.cells.map(cell => (
                                            <td {...cell.getCellProps()}>
                                                {cell.column.id === 'name' || cell.column.id.includes('Theory') || cell.column.id.includes('Practical') ? (
                                                    <input className={`datainput ${cell.column.id == 'name' ? "nameinput" : ""}`}
                                                        value={cell.value || ''}
                                                        onChange={(e) => handleCellEdit(row.original.id, cell.column.id, e.target.value)}
                                                        onBlur={(e) => {
                                                            if (cell.column.id.includes('Theory')) {
                                                                handleCellEdit(row.original.id, cell.column.id, Math.min(Number(e.target.value), 100).toString());
                                                            } else if (cell.column.id.includes('Practical')) {
                                                                handleCellEdit(row.original.id, cell.column.id, Math.min(Number(e.target.value), 100).toString());
                                                            }
                                                        }}
                                                    />
                                                ) : (
                                                    cell.render('Cell')
                                                )}
                                            </td>
                                        ))}
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>
            </div>
            <div className="export-section">
                <button onClick={exportToExcel}>Export to Excel</button>
                <button onClick={exportToImage}>Export to Image</button>
            </div>
            <footer>
                <h3>Designed and developed by <a href="https://github.com/akhil-8605">Akhilesh</a></h3>
            </footer>
        </main>
    );
}
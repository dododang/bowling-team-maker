import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

type TeamData = {
  mainTeams: string[][];
  waitTeams: string[][];
  total: number;
};

export default function App() {
  const [sixData, setSixData] = useState<TeamData | null>(null);
  const [sevenData, setSevenData] = useState<TeamData | null>(null);
  const [dinnerData, setDinnerData] = useState<string[]>([]);
  
  const [rawSixList, setRawSixList] = useState<string[]>([]);
  const [rawSevenList, setRawSevenList] = useState<string[]>([]);
  
  const [executivesInput, setExecutivesInput] = useState<string>("박지헌, 강정은, 김민우, 심영진, 윤지상");

  const shuffleArray = <T,>(array: T[]): T[] => {
    return [...array].sort(() => Math.random() - 0.5);
  };

  const generateTeamsWithExecutives = (list: string[]): TeamData => {
    const execList = executivesInput.split(',').map(name => name.trim()).filter(Boolean);
    const isExecutive = (fullName: string) => execList.some(exec => fullName.includes(exec));

    const execs = shuffleArray(list.filter(isExecutive));
    const regulars = shuffleArray(list.filter(name => !isExecutive(name)));

    const teams: string[][] = [];

    while (execs.length > 0 || regulars.length > 0) {
      const team: string[] = [];
      
      if (execs.length > 0) {
        team.push(execs.pop()!);
      } else {
        team.push(regulars.pop()!);
      }
      
      if (regulars.length > 0) {
        team.push(regulars.pop()!);
      } else if (execs.length > 0) {
        team.push(execs.pop()!);
      }

      teams.push(team);
    }

    const finalTeams = shuffleArray(teams);
    const maxTeams = 12; 
    
    return {
      mainTeams: finalTeams.slice(0, maxTeams),
      waitTeams: finalTeams.slice(maxTeams), // 12팀 초과 인원은 대기로 분리
      total: list.length,
    };
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const arrayBuffer = event.target?.result;
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet) as any[];

      const sixList: string[] = [];
      const sevenList: string[] = [];
      const dinnerList: string[] = [];

      jsonData.forEach((row) => {
        const name = row['이름(*)'] || row['이름'];
        const type = row['볼링 및 회식(*)'] || '';

        if (!name) return;

        if (type.includes('6시')) sixList.push(name);
        if (type.includes('7시')) sevenList.push(name);
        if (type.includes('회식')) dinnerList.push(name);
      });

      setRawSixList(sixList);
      setRawSevenList(sevenList);
      setDinnerData(dinnerList);

      setSixData(generateTeamsWithExecutives(sixList));
      setSevenData(generateTeamsWithExecutives(sevenList));
    };
    reader.readAsArrayBuffer(file);
  };

  const handleReshuffle = () => {
    if (rawSixList.length > 0 && rawSevenList.length > 0) {
      setSixData(generateTeamsWithExecutives(rawSixList));
      setSevenData(generateTeamsWithExecutives(rawSevenList));
    }
  };

  const exportToExcel = async () => {
    if (!sixData || !sevenData) return;

    const execList = executivesInput.split(',').map(name => name.trim()).filter(Boolean);
    const isExecutive = (fullName: string) => execList.some(exec => fullName.includes(exec));

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("조편성_결과");

    ws.columns = [
      { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }
    ];

    const addTimeSection = (title: string, data: TeamData) => {
      if (data.mainTeams.length === 0) return;

      const titleRow = ws.addRow([title, "", "", ""]);
      titleRow.font = { bold: true, size: 14 };
      titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.mergeCells(`A${titleRow.number}:D${titleRow.number}`);

      // ⭐️ 엑셀 표 렌더링을 묶어주는 내부 함수 (정규팀, 대기팀 공통 사용)
      const renderChunk = (teams: string[][], isWait: boolean) => {
        const chunkedTeams = [];
        for (let i = 0; i < teams.length; i += 4) {
          chunkedTeams.push(teams.slice(i, i + 4));
        }

        let offset = 1;
        chunkedTeams.forEach(chunk => {
          const laneHeaders = ["", "", "", ""];
          const members1 = ["", "", "", ""];
          const members2 = ["", "", "", ""];

          chunk.forEach((team, idx) => {
            laneHeaders[idx] = isWait ? `대기 ${offset + idx}팀` : `${offset + idx} 레인`;
            members1[idx] = team[0] || "";
            members2[idx] = team[1] || "";
          });

          const applyStyle = (row: ExcelJS.Row, isHeader: boolean) => {
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              if (colNumber <= 4 && laneHeaders[colNumber - 1] !== "") {
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = {
                  top: { style: 'thin' }, left: { style: 'thin' },
                  bottom: { style: 'thin' }, right: { style: 'thin' }
                };
                if (isHeader) {
                  cell.font = { bold: true, color: isWait ? { argb: 'FF666666' } : undefined };
                  if (isWait) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F5' } };
                }
              }
            });
          };

          const hRow = ws.addRow(laneHeaders);
          applyStyle(hRow, true);

          const m1Row = ws.addRow(members1);
          applyStyle(m1Row, false);
          m1Row.eachCell((cell, colNum) => {
            if (members1[colNum - 1] && isExecutive(members1[colNum - 1])) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } }; 
            }
          });

          const m2Row = ws.addRow(members2);
          applyStyle(m2Row, false);
          m2Row.eachCell((cell, colNum) => {
            if (members2[colNum - 1] && isExecutive(members2[colNum - 1])) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } }; 
            }
          });

          ws.addRow([]); 
          offset += chunk.length;
        });
      };

      // 정규 팀 엑셀 출력
      renderChunk(data.mainTeams, false);

      // ⭐️ 대기 팀 엑셀 출력 (누락되었던 부분)
      if (data.waitTeams.length > 0) {
        const waitTitleRow = ws.addRow(["[ 대기 팀 편성 ]", "", "", ""]);
        waitTitleRow.font = { bold: true, color: { argb: 'FF555555' } };
        ws.mergeCells(`A${waitTitleRow.number}:D${waitTitleRow.number}`);
        renderChunk(data.waitTeams, true);
      }

      ws.addRow([]); 
    };

    addTimeSection("6시", sixData);
    addTimeSection("7시", sevenData);

    if (dinnerData.length > 0) {
      const dinnerTitleRow = ws.addRow(["회식 참여 인원", "", "", ""]);
      dinnerTitleRow.font = { bold: true, size: 14 };
      dinnerTitleRow.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.mergeCells(`A${dinnerTitleRow.number}:D${dinnerTitleRow.number}`);
      
      const chunkedDinner = [];
      for (let i = 0; i < dinnerData.length; i += 4) {
        chunkedDinner.push(dinnerData.slice(i, i + 4));
      }
      
      chunkedDinner.forEach(chunk => {
         const row = ws.addRow([chunk[0] || "", chunk[1] || "", chunk[2] || "", chunk[3] || ""]);
         row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
           if (colNumber <= 4 && (chunk[0] || chunk[1] || chunk[2] || chunk[3])) {
              cell.alignment = { horizontal: 'center', vertical: 'middle' };
              cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
              };
           }
         });
      });
    }

    const buffer = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), "볼링조편성_결과.xlsx");
  };

  const renderTeamSection = (title: string, data: TeamData) => (
    <div className="mb-8 p-6 bg-white rounded-lg shadow-md border-t-4 border-blue-500">
      <h2 className="text-2xl font-bold mb-4 text-blue-700">
        🎳 {title} 타임 (총 {data.total}명)
      </h2>
      <h3 className="font-bold text-gray-800 mb-3 text-lg border-b pb-2">
        [ 정규 팀 편성 - 총 {data.mainTeams.length}팀 ]
      </h3>
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4 mb-6">
        {data.mainTeams.map((team, idx) => (
          <div key={idx} className="p-4 bg-blue-50 hover:bg-blue-100 transition border border-blue-200 rounded-lg shadow-sm">
            <span className="font-extrabold text-blue-600 block mb-1">{idx + 1}팀</span>
            <span className="text-gray-700 font-medium">
              {team.map((member, i) => (
                <span key={i} className={executivesInput.includes(member.split(' ')[0]) ? "bg-green-200 px-1 rounded mr-1" : "mr-1"}>
                  {member}
                </span>
              ))}
            </span>
          </div>
        ))}
      </div>

      {/* ⭐️ 누락되었던 대기 팀 웹 화면 출력 복구 */}
      {data.waitTeams.length > 0 && (
        <>
          <h3 className="font-bold text-gray-500 mb-3 text-lg border-b pb-2 mt-8">
            [ 대기 팀 편성 - 총 {data.waitTeams.length}팀 ]
          </h3>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4 mb-6">
            {data.waitTeams.map((team, idx) => (
              <div key={idx} className="p-4 bg-gray-50 border border-gray-200 rounded-lg shadow-sm">
                <span className="font-bold text-gray-500 block mb-1">대기 {idx + 1}팀</span>
                <span className="text-gray-600 font-medium">
                  {team.map((member, i) => (
                    <span key={i} className={executivesInput.includes(member.split(' ')[0]) ? "bg-green-200 px-1 rounded mr-1" : "mr-1"}>
                      {member}
                    </span>
                  ))}
                </span>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );

  return (
    <div className="min-h-screen bg-gray-100 p-8 font-sans">
      <div className="max-w-5xl mx-auto">
        <h1 className="text-4xl font-black mb-10 text-center text-gray-800 tracking-tight">
          볼링 대회 조 편성 자동화 🎲
        </h1>

        <div className="bg-white p-6 rounded-lg shadow-md mb-8 border border-gray-200">
          <label className="block text-gray-800 font-bold mb-2 text-lg">
            👑 임원진 명단 (쉼표로 구분하여 입력)
          </label>
          <input 
            type="text" 
            value={executivesInput}
            onChange={(e) => setExecutivesInput(e.target.value)}
            className="w-full p-3 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="예: 박지헌, 강정은, 김민우, 심영진, 윤지상"
          />
          <p className="text-sm text-gray-500 mt-2">
            * 이곳에 입력된 임원진은 <b>서로 같은 팀이 되지 않도록 1명씩 분산 배치</b>됩니다.
          </p>
        </div>
        
        <div className="flex justify-center flex-wrap gap-4 mb-12">
          <label className="cursor-pointer bg-gray-900 text-white px-8 py-4 rounded-xl font-bold text-lg shadow-lg hover:bg-gray-700 hover:scale-105 transition-all duration-200">
            엑셀 파일 업로드 (.xlsx)
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              onChange={handleFileUpload} 
              className="hidden" 
            />
          </label>

          {sixData && sevenData && (
            <>
              <button 
                onClick={handleReshuffle}
                className="bg-blue-600 text-white px-8 py-4 rounded-xl font-bold text-lg shadow-lg hover:bg-blue-500 hover:scale-105 transition-all duration-200"
              >
                조 편성 다시 섞기 🎲
              </button>
              
              <button 
                onClick={exportToExcel}
                className="bg-green-600 text-white px-8 py-4 rounded-xl font-bold text-lg shadow-lg hover:bg-green-500 hover:scale-105 transition-all duration-200"
              >
                결과 엑셀로 다운로드 📥
              </button>
            </>
          )}
        </div>

        {sixData && sevenData && (
          <div className="space-y-8 animate-fade-in-up">
            {renderTeamSection('6시', sixData)}
            {renderTeamSection('7시', sevenData)}

            <div className="p-6 bg-white rounded-lg shadow-md border-t-4 border-orange-400">
              <h2 className="text-2xl font-bold mb-4 text-orange-600">
                🍻 회식 참여 인원 (총 {dinnerData.length}명)
              </h2>
              <div className="p-5 bg-orange-50 border border-orange-100 rounded-lg text-gray-800 font-medium leading-relaxed shadow-sm">
                {dinnerData.length > 0 ? dinnerData.join(', ') : '참여 인원이 없습니다.'}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

import React, { useState, useCallback, useRef } from 'react';
import { GoogleGenAI } from "@google/genai";
import * as XLSX from 'xlsx';
import type { RowData } from './types';
import { UploadIcon, SparklesIcon, DownloadIcon, CheckCircleIcon, XCircleIcon } from './components/icons';

const ProgressBar: React.FC<{ progress: number }> = ({ progress }) => (
    <div className="w-full bg-slate-700 rounded-full h-2.5">
        <div
            className="bg-purple-600 h-2.5 rounded-full transition-all duration-300 ease-linear"
            style={{ width: `${progress}%` }}
        ></div>
    </div>
);


const DiffHighlight: React.FC<{ original: string; corrected: string; mode: 'original' | 'corrected' }> = ({ original, corrected, mode }) => {
    const safeOriginal = original || '';
    const safeCorrected = corrected || '';

    if (safeOriginal === safeCorrected) {
        return <>{safeOriginal}</>;
    }

    const originalTokens = safeOriginal.split(/(\s+)/);
    const correctedTokens = safeCorrected.split(/(\s+)/);

    // If token counts differ, use a more robust set-based comparison for highlighting.
    if (originalTokens.length !== correctedTokens.length) {
        const correctedWords = new Set(correctedTokens.filter(t => t.trim() !== ''));
        const originalWords = new Set(originalTokens.filter(t => t.trim() !== ''));

        const tokensToRender = mode === 'original' ? originalTokens : correctedTokens;
        const wordsToCompare = mode === 'original' ? correctedWords : originalWords;
        const highlightClass = mode === 'original' 
            ? 'text-red-400 bg-red-900/30 rounded px-1' 
            : 'text-green-300 bg-green-900/30 rounded px-1';
        
        return (
            <>
                {tokensToRender.map((token, index) => {
                    // Highlight if it's a non-whitespace word that doesn't exist in the other set.
                    if (token.trim() !== '' && !wordsToCompare.has(token)) {
                        return <span key={`${index}-${token}`} className={highlightClass}>{token}</span>;
                    }
                    return <React.Fragment key={`${index}-${token}`}>{token}</React.Fragment>;
                })}
            </>
        );
    }
    
    // Original logic for when token counts are the same.
    const displayTokens = mode === 'original' ? originalTokens : correctedTokens;

    return (
        <>
            {originalTokens.map((token, index) => {
                const correctedToken = correctedTokens[index];
                if (token !== correctedToken) {
                    const style = mode === 'original'
                        ? 'text-red-400 bg-red-900/30 rounded px-1'
                        : 'text-green-300 bg-green-900/30 rounded px-1';
                    return <span key={`${index}-${token}`} className={style}>{displayTokens[index]}</span>;
                }
                return <React.Fragment key={`${index}-${token}`}>{token}</React.Fragment>;
            })}
        </>
    );
};


const SpellingResultsTable: React.FC<{ originalData: RowData[], correctedData: RowData[] }> = ({ originalData, correctedData }) => {
    const findCorrectedRow = (id: number) => correctedData.find(row => row.id === id);

    const changedRows = originalData.filter(originalRow => {
        const correctedRow = findCorrectedRow(originalRow.id);
        if (!correctedRow) return false;
        const storyChanged = (originalRow['story'] || '') !== (correctedRow['story'] || '');
        const subStoryChanged = (originalRow['sub-story'] || '') !== (correctedRow['sub-story'] || '');
        return storyChanged || subStoryChanged;
    });

    return (
        <div className="mt-8 w-full overflow-x-auto">
            {changedRows.length > 0 ? (
                <table className="min-w-full bg-slate-800 border border-slate-700 rounded-lg shadow-lg">
                    <thead className="bg-slate-700/50">
                        <tr>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-slate-400 uppercase tracking-wider">Original Story</th>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-slate-400 uppercase tracking-wider">Corrected Story</th>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-slate-400 uppercase tracking-wider">Original Sub-Story</th>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-slate-400 uppercase tracking-wider">Corrected Sub-Story</th>
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-700">
                        {changedRows.map(originalRow => {
                            const correctedRow = findCorrectedRow(originalRow.id);
                            if (!correctedRow) return null;

                            const originalStory = originalRow['story'] || '';
                            const correctedStory = correctedRow['story'] || '';
                            const originalSubStory = originalRow['sub-story'] || '';
                            const correctedSubStory = correctedRow['sub-story'] || '';

                            return (
                                <tr key={originalRow.id} className="hover:bg-slate-700/40 transition-colors duration-200">
                                    <td className="px-4 py-3 border-b border-slate-700 text-slate-400 align-top">
                                        <DiffHighlight original={originalStory} corrected={correctedStory} mode="original" />
                                    </td>
                                    <td className="px-4 py-3 border-b border-slate-700 text-slate-300 align-top">
                                        <DiffHighlight original={originalStory} corrected={correctedStory} mode="corrected" />
                                    </td>
                                    <td className="px-4 py-3 border-b border-slate-700 text-slate-400 align-top">
                                        <DiffHighlight original={originalSubStory} corrected={correctedSubStory} mode="original" />
                                    </td>
                                    <td className="px-4 py-3 border-b border-slate-700 text-slate-300 align-top">
                                        <DiffHighlight original={originalSubStory} corrected={correctedSubStory} mode="corrected" />
                                    </td>
                                </tr>
                            );
                        })}
                    </tbody>
                </table>
            ) : (
                <div className="text-center p-8 bg-slate-800 border border-slate-700 rounded-lg">
                    <p className="text-slate-300">No spelling corrections were found in the uploaded file.</p>
                </div>
            )}
        </div>
    );
};

const Spinner: React.FC = () => (
    <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
);


export default function App() {
    const [file, setFile] = useState<File | null>(null);
    const [fileName, setFileName] = useState<string>('');
    const [originalData, setOriginalData] = useState<RowData[]>([]);
    const [correctedData, setCorrectedData] = useState<RowData[]>([]);
    const [isProcessing, setIsProcessing] = useState<boolean>(false);
    const [error, setError] = useState<string>('');
    const [progressMessage, setProgressMessage] = useState<string>('');
    const [processingProgress, setProcessingProgress] = useState<number>(0);
    
    const fileInputRef = useRef<HTMLInputElement>(null);

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const selectedFile = e.target.files?.[0];
        if (selectedFile) {
            if (selectedFile.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' && selectedFile.type !== 'application/vnd.ms-excel') {
                 setError('Invalid file type. Please upload an Excel file (.xlsx, .xls).');
                 return;
            }
            setFile(selectedFile);
            setFileName(selectedFile.name);
            setError('');
            setCorrectedData([]);
            setOriginalData([]);
            setProgressMessage('File ready. Click "Correct Spelling" to process.');

            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target?.result as ArrayBuffer);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet);

                    const lowercaseHeaders = (row: any) => 
                        Object.keys(row).reduce((acc, key) => {
                            acc[key.toLowerCase()] = row[key];
                            return acc;
                        }, {} as {[key: string]: any});

                    const processedData = jsonData.map(lowercaseHeaders);

                    if (!processedData[0]?.hasOwnProperty('story') || !processedData[0]?.hasOwnProperty('sub-story')) {
                        setError("Excel sheet must contain 'story' and 'sub-story' columns.");
                        setOriginalData([]);
                        return;
                    }

                    setOriginalData(processedData.map((row, index) => ({ id: index, ...row })));
                } catch (err) {
                    setError('Error reading Excel file. Please ensure it is a valid format.');
                    console.error(err);
                }
            };
            reader.readAsArrayBuffer(selectedFile);
        }
    };
    
    const triggerFileSelect = () => fileInputRef.current?.click();

    const handleProcess = useCallback(async () => {
        if (!originalData.length) {
            setError('No data to process. Please upload a valid Excel file.');
            return;
        }

        setIsProcessing(true);
        setError('');
        setProgressMessage('Initializing analysis...');
        setProcessingProgress(0);
        setCorrectedData([]);

        try {
            setProgressMessage('Identifying unique rows...');
            const uniqueDataMap = new Map<string, { id: number; story: string; 'sub-story': string }>();
            const uniqueKeyToOriginalIdsMap = new Map<string, number[]>();
            let uniqueIdCounter = 0;

            originalData.forEach(row => {
                const story = row['story'] || "";
                const subStory = row['sub-story'] || "";
                const key = `${story}|~|${subStory}`;
                
                if (!uniqueDataMap.has(key)) {
                    uniqueDataMap.set(key, {
                        id: uniqueIdCounter++,
                        story: story,
                        'sub-story': subStory,
                    });
                    uniqueKeyToOriginalIdsMap.set(key, []);
                }
                uniqueKeyToOriginalIdsMap.get(key)!.push(row.id);
            });
            
            const dataToAnalyze = Array.from(uniqueDataMap.values());
            setProgressMessage(`Found ${dataToAnalyze.length} unique rows to process.`);

            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY as string });

            const CHUNK_SIZE = 20;
            const DELAY_MS = 1500; // 1.5-second delay to stay safely within API rate limits.
            const delay = (ms: number) => new Promise(res => setTimeout(res, ms));

            let cumulativeResults: any[] = [];

            for (let i = 0; i < dataToAnalyze.length; i += CHUNK_SIZE) {
                const chunk = dataToAnalyze.slice(i, i + CHUNK_SIZE);
                const progress = (i / dataToAnalyze.length) * 100;
                setProcessingProgress(progress);
                setProgressMessage(`Processing unique rows ${i + 1} to ${Math.min(i + CHUNK_SIZE, dataToAnalyze.length)} of ${dataToAnalyze.length}...`);

                const prompt = `
                    You are an AI proofreader with a single, precise task: correct spelling mistakes in the 'story' and 'sub-story' fields of the provided JSON array. Follow these rules strictly.
                    
                    **CRITICAL RULES:**
                    1.  **SPELLING ONLY:** Correct only obvious spelling errors.
                    2.  **NO GRAMMAR/PUNCTUATION:** Do NOT change grammar. Do NOT add, remove, or alter punctuation. Specifically, DO NOT add apostrophes. For example, "SINHAS" should remain "SINHAS", not "SINHA'S". "BJP S" should remain "BJP S", not "BJP'S".
                    3.  **PRESERVE ALL CONTENT:** Do not change names, numbers, acronyms, or the meaning of the text. Do not add or remove words. For example, do not change 'ADIMINISTRATION' to 'ADITI MISHRA'. Correct it to 'ADMINISTRATION' if that's the clear intent.
                    4.  **HANDLE PROPER NOUNS CAREFULLY:** Correct obvious misspellings of proper nouns. For example, 'UTTARAKHANDA' should be corrected to 'UTTARAKHAND'.
                    5.  **EXACT JSON STRUCTURE:** The output MUST be a single, valid, minified JSON array. It must have the exact same number of objects and 'id's as the input. Do not wrap the JSON in markdown.
                    6.  **IF NO ERRORS, NO CHANGE:** If a field has no spelling errors, return it exactly as it is.

                    Input Data:
                    ${JSON.stringify(chunk)}
                `;

                const response = await ai.models.generateContent({
                    model: "gemini-flash-lite-latest",
                    contents: prompt,
                });

                const responseText = response.text;
                if (!responseText) {
                    throw new Error(`Model returned an empty response for the chunk starting at row ${i + 1}. The response might have been blocked.`);
                }

                const cleanedResponse = responseText.replace(/```json|```/g, '').trim();
                const resultChunk = JSON.parse(cleanedResponse);

                 if (Array.isArray(resultChunk)) {
                    cumulativeResults.push(...resultChunk);
                } else {
                     throw new Error(`Model returned invalid data for a chunk starting at row ${i+1}.`);
                }
                
                if (i + CHUNK_SIZE < dataToAnalyze.length) {
                    await delay(DELAY_MS);
                }
            }
            
            setProcessingProgress(100);
            setProgressMessage('Mapping results back to original rows...');
            
            const expandedResults: any[] = [];
            const uniqueResultsMap = new Map(cumulativeResults.map(res => [res.id, res]));
    
            uniqueDataMap.forEach((uniqueRowData, key) => {
                const originalIds = uniqueKeyToOriginalIdsMap.get(key);
                const result = uniqueResultsMap.get(uniqueRowData.id);
    
                if (originalIds && result) {
                    originalIds.forEach(originalId => {
                        expandedResults.push({
                            id: originalId,
                            story: result.story,
                            'sub-story': result['sub-story'],
                        });
                    });
                }
            });

            setCorrectedData(expandedResults);
            setProgressMessage(`Correction complete. ${originalData.length} rows processed based on ${dataToAnalyze.length} unique entries.`);

        } catch (err: any) {
            console.error(err);
            setError(`An error occurred: ${err.message}. The model may have returned an invalid format. Please try again.`);
            setProgressMessage('');
            setProcessingProgress(0);
        } finally {
            setIsProcessing(false);
        }
    }, [originalData]);

    const handleDownload = () => {
        if (correctedData.length > 0) {
            const finalData = originalData.map(originalRow => {
                const correctedRow = correctedData.find(cr => cr.id === originalRow.id);
                const newRow: { [key: string]: any } = {};
    
                Object.keys(originalRow).forEach(key => {
                    newRow[key] = originalRow[key];
                    if (key === 'story') {
                        newRow['corrected_story'] = correctedRow ? correctedRow['story'] : originalRow['story'];
                    }
                    if (key === 'sub-story') {
                        newRow['corrected_sub-story'] = correctedRow ? correctedRow['sub-story'] : originalRow['sub-story'];
                    }
                });
                return newRow;
            });
            const dataToExport = finalData.map(({ id, ...rest }) => rest);
            const worksheet = XLSX.utils.json_to_sheet(dataToExport);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Corrected_Data");
            XLSX.writeFile(workbook, "corrected_spelling.xlsx");
        }
    };

    const hasResults = correctedData.length > 0;

    return (
        <div className="min-h-screen bg-slate-900 text-slate-100 flex flex-col items-center p-4 sm:p-6 md:p-8 font-sans">
            <div className="w-full max-w-7xl mx-auto">
                <header className="text-center mb-8">
                    <h1 className="text-4xl sm:text-5xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-purple-400 to-cyan-400">
                        Excel Spelling Corrector
                    </h1>
                    <p className="mt-4 text-lg text-slate-400 max-w-3xl mx-auto">
                        Upload your Excel file to automatically correct spelling mistakes in the 'story' and 'sub-story' columns.
                    </p>
                </header>

                <main className="bg-slate-800/50 border border-slate-700 rounded-xl shadow-2xl p-6 md:p-8 flex flex-col items-center gap-6">
                    <div className="w-full max-w-2xl flex flex-col gap-4 items-center">
                        <div className="w-full flex flex-col sm:flex-row gap-4">
                            <input
                                type="file"
                                ref={fileInputRef}
                                onChange={handleFileChange}
                                className="hidden"
                                accept=".xlsx, .xls"
                            />
                            <button
                                onClick={triggerFileSelect}
                                className="w-full sm:w-auto flex-grow flex items-center justify-center gap-2 px-6 py-3 bg-slate-700 text-slate-200 rounded-lg hover:bg-slate-600 transition-all duration-300 focus:outline-none focus:ring-2 focus:ring-purple-500 focus:ring-offset-2 focus:ring-offset-slate-900"
                            >
                                <UploadIcon className="w-5 h-5" />
                                {fileName ? 'Change File' : 'Select Excel File'}
                            </button>

                            <button
                                onClick={handleProcess}
                                disabled={!file || isProcessing || originalData.length === 0}
                                className="w-full sm:w-auto flex items-center justify-center gap-2 px-6 py-3 bg-purple-600 text-white font-semibold rounded-lg shadow-md hover:bg-purple-700 transition-all duration-300 disabled:bg-slate-500 disabled:cursor-not-allowed focus:outline-none focus:ring-2 focus:ring-purple-500 focus:ring-offset-2 focus:ring-offset-slate-900"
                            >
                                {isProcessing ? <Spinner/> : <SparklesIcon className="w-5 h-5" />}
                                {isProcessing ? 'Correcting...' : 'Correct Spelling'}
                            </button>
                        </div>
                    </div>

                    <div className="text-center h-12 w-full max-w-md flex flex-col justify-center items-center">
                         {isProcessing ? (
                            <div className="w-full">
                                <p className="text-cyan-400 mb-2">{progressMessage}</p>
                                <ProgressBar progress={processingProgress} />
                            </div>
                        ) : error ? (
                             <div className="flex items-center gap-2 text-red-400">
                               <XCircleIcon className="w-5 h-5" /> <span>{error}</span>
                            </div>
                        ) : progressMessage && fileName && !hasResults ? (
                             <div className="flex items-center gap-2 text-green-400">
                               <CheckCircleIcon className="w-5 h-5" /> <span>{progressMessage}</span>
                            </div>
                        ) : null}
                    </div>
                    
                    {hasResults && !isProcessing && (
                         <div className="w-full flex flex-col items-center gap-6">
                             <div className="flex items-center gap-2 text-green-400">
                               <CheckCircleIcon className="w-5 h-5" /> <span>{progressMessage}</span>
                            </div>
                            <button
                                onClick={handleDownload}
                                className="flex items-center gap-2 px-6 py-3 bg-green-600 text-white font-semibold rounded-lg shadow-md hover:bg-green-700 transition-all duration-300 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 focus:ring-offset-slate-900"
                            >
                                <DownloadIcon className="w-5 h-5" />
                                Download Corrected Excel
                            </button>
                            <SpellingResultsTable originalData={originalData} correctedData={correctedData} />
                        </div>
                    )}
                </main>
            </div>
        </div>
    );
}
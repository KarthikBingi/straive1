// json2html.js

export default function json2html(data) {
    // Extract unique keys for the table headers
    const headers = [...new Set(data.flatMap(Object.keys))];
    
    // Begin the HTML string with the table and data-user attribute
    let html = '<table data-user="bingikarthik2014@gmail.com"><thead><tr>';
    
    // Generate table headers
    html += headers.map(header => `<th>${header}</th>`).join('');
    html += '</tr></thead><tbody>';
    
    // Generate table rows
    html += data.map(row => {
        return `<tr>${headers.map(header => `<td>${row[header] || ''}</td>`).join('')}</tr>`;
    }).join('');
    
    html += '</tbody></table>';
    return html;
}

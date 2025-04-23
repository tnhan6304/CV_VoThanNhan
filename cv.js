document.getElementById("exportWord").addEventListener("click", function () {
    const content = document.querySelector(".container").innerHTML;
    const styles = Array.from(document.querySelectorAll('link[rel="stylesheet"], style'))
        .map(style => {
            if (style.tagName === "LINK") {
                const xhr = new XMLHttpRequest();
                xhr.open("GET", style.href, false); 
                xhr.send();
                return `<style>${xhr.responseText}</style>`;
            } else {
                return `<style>${style.innerHTML}</style>`;
            }
        })
        .join("");
    const header = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="utf-8"><title>Export Word</title>${styles}</head><body>`;
    const footer = `</body></html>`;
    const sourceHTML = header + content + footer;
    const blob = new Blob(['\ufeff', sourceHTML], {
        type: 'application/msword'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "CV_VoThanhNhan.doc";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});
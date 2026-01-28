function search() {
    const query = document.getElementById('search').value;
    const resultsDiv = document.getElementById('results');
    if (query.trim() === '') {
        resultsDiv.innerHTML = 'Te rog introdu un termen de căutare.';
    } else {
        resultsDiv.innerHTML = `Rezultatele pentru "${query}": Aceasta este o căutare simplă. Într-o aplicație reală, aici ar fi rezultatele.`;
    }
}
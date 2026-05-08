# Copilot Instructions - lab01

## Project Overview
This is a dual-interface project showcasing the same core functionality implemented two ways:
- **Web-based**: HTML/CSS/JavaScript frontend (index.html, style.css, script.js)
- **Desktop GUI**: Python Tkinter application (gui_app.py)

Both interfaces provide identical features: About section, Contact form, and Search functionality. Content is in Romanian.

## Architecture & Key Patterns

### UI Layer Structure
- **Web**: Single-page layout with section-based navigation (Sobre, Contact, Caută)
- **Desktop**: Tab-based interface using ttk.Notebook with three tabs matching web sections
- Both implement synchronous, simple placeholder functionality without backend data persistence

### Common UI Flows
1. **About/Despre**: Display static information
2. **Contact**: Form collection (Nume, Email, Mesaj) → messagebox confirmation (Python) or div display (Web)
3. **Search/Caută**: Simple query validation → placeholder results message

### Language & Localization
- All UI text and labels use Romanian language
- Variable names follow Romanian conventions (e.g., `entry_nume`, `text_mesaj`, `resultsDiv`)
- Maintain this language consistency when modifying UI strings

## Important Implementation Notes

### Python GUI (gui_app.py)
- Uses tkinter with ttk.Notebook for tabbed interface
- Functions (`show_about()`, `submit_contact()`, `search()`) handle button clicks
- Messagebox dialogs provide user feedback (showinfo, showwarning)
- Entry widgets capture form input; text widget used for multi-line message
- Window size: 400x400, simple grid layout without explicit layout managers

### Web Frontend (index.html, script.js, style.css)
- Vanilla JavaScript (no frameworks) - keep it minimal and dependency-free
- Form submission in HTML structure; JavaScript only enhances with `search()` function
- Script dynamically updates `#results` div with query results
- CSS provides header navigation styling, basic form layout
- Mobile-friendly: includes viewport meta tag

## Developer Workflows

### Running the Application
- **Python GUI**: `python gui_app.py` - requires tkinter (built-in for most Python installations)
- **Web**: Open `index.html` in a browser directly (file:// protocol works)

### Testing
- Manual testing only currently (test.md exists but contains minimal content)
- Test both Romanian text rendering and form functionality
- Verify validation: search requires non-empty query

## Critical Conventions to Maintain

1. **Parallel Implementation**: Keep Python and JavaScript versions feature-identical
2. **Romanian UI Text**: All user-facing strings remain in Romanian
3. **No External Dependencies**: Web uses pure JavaScript; Python uses only tkinter
4. **Placeholder Semantics**: Search/Contact functions show demo messages, not real data
5. **Tab/Section Naming**: Always match the three sections (Despre, Contact, Caută)

## When Making Changes
- Update both implementations (Python AND Web) for feature parity
- Test form input validation on both platforms
- Preserve Romanian language in all UI labels
- Keep the simple, educational nature - no complex state management

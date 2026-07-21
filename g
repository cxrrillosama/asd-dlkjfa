1. Ditch the "Card" Look
The current design puts every macro inside its own bordered box. Cards require padding on the left, right, top, and bottom, which eats up valuable horizontal space.

Edge-to-Edge List Items: Instead of individual boxes, make the macros a flat, edge-to-edge list. Use a very subtle alternating background color (zebra striping) or just a simple, single-pixel separator line between items.

Result: This gives you back 10–20 pixels of horizontal space, preventing longer macro names from wrapping to a second line.

2. Condense the Filters
Right now, the filters (Ticket Summary, Welcome, Work Notes) are displayed as pill-shaped buttons. In a narrow sidebar, these will quickly wrap to multiple lines, pushing the actual macros further down the page.

Use a Dropdown: Replace the horizontal row of pills with a single dropdown menu labeled "Filter by tag..." or a funnel icon.

Segmented Control: If there are strictly only ever two or three categories, use a compact segmented control (like a toggle switch but for three items) right below the search bar.

3. Right-Align Action Icons
In the current screenshot, the favorite star sits on the left, underneath the text. This forces the eye to jump around and takes up vertical space.

Flush Right, Reveal on Hover: Keep the macro text firmly left-aligned for easy reading. Put any action icons (like a star for favorites, or three dots for a context menu) flush against the right edge of the sidebar.

Keep it Clean: Hide these icons until the user hovers over the specific macro row. This keeps the narrow column from looking cluttered.

4. Leverage Accordions Heavily
Since you can't expand horizontally, you have to be extremely smart about vertical scrolling.

Start Collapsed: If a user has dozens of macros across multiple folders, default to showing the folders collapsed.

Sticky Headers: When a user expands a folder (e.g., "Work Notes") and scrolls down, make that folder title "stick" to the top of the sidebar under the search bar so they never lose context of where they are in the list.

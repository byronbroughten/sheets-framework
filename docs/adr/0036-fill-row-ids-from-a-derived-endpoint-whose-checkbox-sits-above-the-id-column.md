# Fill row IDs from a derived endpoint whose checkbox sits above the ID column

A person fills a Table's blank row IDs by ticking a Fill Row IDs checkbox, which fires an endpoint the framework derives for every Table with an ID column. It is an endpoint, not a menu item, because a menu runs as whoever clicks it, so every editor would have to approve the script's permissions, while an endpoint runs as the account that installed its trigger. The checkbox sits in the action-row cell above the ID column, because ID sits right of the frozen Name columns, and the Table name chip hides only the first column's action-row cell.

## Considered options

- **A custom menu item**: rejected for the permission prompt every editor would meet.
- **The checkbox above the Table's first column**: rejected, since the name chip covers that cell; a Table whose ID is still first gets the `idFirst` notice instead.

Source: byronbroughten/sheets-framework#219, byronbroughten/sheets-framework#248

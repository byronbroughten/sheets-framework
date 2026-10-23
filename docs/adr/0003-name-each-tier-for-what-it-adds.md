# Name each tier for what it adds: Source, Schema, Raw, Identified, Named

Source knows nothing about this spreadsheet, Schema is the declared layout plus everything that reads the generated configs, Raw addresses by position and is blind to column config, Identified adds the committed column config reached by column ID, and Named adds names. Schema sits below Raw because Raw reads the layout; injecting the layout into Raw as props was rejected, since Raw has no use without one. Indexed was renamed Identified so that "index" means position everywhere, and a lint rule, not a convention, keeps column config out of Raw.

Source: byronbroughten/byro-repo#71, byronbroughten/byro-repo#72, byronbroughten/sheets-framework#127

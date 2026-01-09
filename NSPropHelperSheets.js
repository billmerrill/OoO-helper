const NSPropertyHelperSheets = {
  fauxProps: {},
  NS: 'OoO-',

  /**
   * Helper to get the Document object safely
   */
  _getDoc: function() {
    console.log("hi there")
    return DocumentApp.getActiveDocument();
  },

  get: function(key) {
    try {
      const metadata = this._getDoc().getDeveloperMetadata();
      console.log('starting get');
      console.log(metadata);
      const match = metadata.find(m => m.getKey() === this.NS + key);
      return match ? match.getValue() : undefined;
    } catch(e) {
      console.warn("DeveloperMetadata unavailable, using js mock.");
      return this.fauxProps[this.NS + key] || undefined;
    }
  },

  getAll: function() {
    let original = {};
    try {
      const metadata = this._getDoc().getDeveloperMetadata();
      metadata.forEach(m => {
        original[m.getKey()] = m.getValue();
      });
    } catch(e) {
      console.warn("DeveloperMetadata unavailable, using js mock.");
      original = this.fauxProps;
    }

    const cleaned = Object.entries(original).reduce(
      (acc, [key, value]) => {
        if (key.startsWith(this.NS)) {
          const newKey = key.replace(this.NS, "");
          acc[newKey] = value;
        }
        return acc;
      }, {});

    return cleaned;
  },

  set: function(key, value) {
    const fullKey = this.NS + key;
    try {
      // 1. Remove existing instances of this key to mimic "update" behavior
      this.delete(key); 
      
      // 2. Add the new metadata
      this._getDoc().addDeveloperMetadata(fullKey, value);
    } catch (e) {
      console.warn("DeveloperMetadata unavailable, using js mock.");
      this.fauxProps[fullKey] = value;
    }
  },

  delete: function(key) {
    const fullKey = this.NS + key;
    try {
      const metadata = this._getDoc().getDeveloperMetadata();
      metadata.forEach(m => {
        if (m.getKey() === fullKey) {
          m.remove();
        }
      });
    } catch (e) {
      console.warn("DeveloperMetadata unavailable, using js mock.");
      delete this.fauxProps[fullKey];
    }
  },
};
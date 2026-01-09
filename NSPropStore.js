const NSPropertyHelper = {
  fauxProps: {},
  NS: 'OoO-',

  get: function(key) {
    try {
      return PropertiesService.getDocumentProperties().getProperty(this.NS + key);
    } catch(e) {
      console.warn(e);
      console.warn("DocumentProperties unavailable, using js mock.");
      return this.fauxProps[this.NS + key] || undefined;
    }
  },

  getAll: function() {
    var original;
    try {
      original = PropertiesService.getDocumentProperties().getProperties();
    } catch(e) {
      console.warn(e);
      console.warn("DocumentProperties unavailable, using js mock.");
      original = this.fauxProps;
    }

    const cleaned = Object.entries(original).reduce(
      (acc, [key, value]) => {
        const newKey = key.startsWith(this.NS) ? key.replace(this.NS, "") : key;
        acc[newKey] = value;
        return acc;
      }, {})

    return cleaned;
  },

  set: function(key, value) {

    try {
      const docProps = PropertiesService.getDocumentProperties()
      docProps.setProperty(this.NS + key, value);
    } catch (e) {
      console.warn(e);
      console.warn("DocumentProperties unavailable, using js mock.");
      this.fauxProps[this.NS + key] = value;
    }
  },

  /**
     * Accepts an object of key-value pairs and saves them in bulk.
     * @param {Object} dataObject - e.g., { "userId": "123", "theme": "dark" }
     */
  setAll: function(dataObject) {
    // 1. Create a new object with the namespace applied to all keys
    const namespacedData = {};
    for (let key in dataObject) {
      if (dataObject.hasOwnProperty(key)) {
        namespacedData[this.NS + key] = dataObject[key];
      }
    }

    try {
      const docProps = PropertiesService.getDocumentProperties();
      // 2. Use setProperties for a single batch write. 
      // The 'false' argument ensures we don't delete existing properties.
      docProps.setProperties(namespacedData, false);
    } catch (e) {
      console.warn(e);
      console.warn("DocumentProperties unavailable, saving all to js mock.");
      
      // Fallback to mock object
      for (let nsKey in namespacedData) {
        this.fauxProps[nsKey] = namespacedData[nsKey];
      }
    }
  },


  delete: function(key) {
    try {
      return PropertiesService.getDocumentProperties().deleteProperty(this.NS + key);
    } catch (e) {
      console.warn("DocumentProperties unavailable, using js mock.");
      delete this.fauxProps[this.NS + key];
    }
  },
};

function getSettings() {
  return NSPropertyHelper.getAll();
}

function saveSettings(settings) {
  const saveSettings = {'eventTitle': settings['eventTitle'],
                         'p1Name': settings['p1Name'],
                         'p2Name': settings['p2Name']}
  NSPropertyHelper.setAll(saveSettings);
}


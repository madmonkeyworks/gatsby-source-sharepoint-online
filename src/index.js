const { createClient } = require("./client");
const { Resource } = require("./resource");

exports.onPreInit = () =>
  console.log("Loaded gatsby-source-sharepoint-online plugin");

/**
 * @typedef {import('gatsby').SourceNodesArgs} Helpers
 */

/**
 * Generates Gatsby source nodes attached to a Sharepoint Online tenant.
 * @param {Helpers} helpers Gatsby Node Helpers.
 * @param {any} config Config object provided in the plugin config.
 */
exports.sourceNodes = async (helpers, config) => {
  const { host, sites = [], ...creds } = config;
  const client = createClient(creds);
  
  // LISTS
  const listResource = new Resource("list");
  for (let i = 0; i < sites.length; i++) {
    const { lists = [] } = sites[i];

    const get = listResource.requestLists(host, sites[i], client, helpers);

    for (let j = 0; j < lists.length; j++) {
      const list = lists[j];

      if (!listResource.validateList(list)) {
        continue;
      }

      try {
        await get(list);
      } catch (err) {
        console.error(err);
      }
    }
  }
};

/**
 * Creates custom schema for Sharepoint image fields.
 * @param {Helpers} actions Gatsby Node Helpers.
 */
exports.createSchemaCustomization = ({ actions }, config) => {
  const { createTypes } = actions;
  const { host, sites = [], ...creds } = config;
  const listResource = new Resource("list");

  sites.forEach(site => {
    const { lists = [] } = site;

    lists.forEach(list => {
      
      if (!listResource.validateList(list)) {
        return
      }

      const nodeType = listResource.generateListNodeType(list.title)
      
      if ( typeof list.fields !== undefined ) {
        const imageFields = list.fields.filter(f => f?.fieldType === "image").map(f => f.fieldName);
        imageFields.forEach(field => {
          const imageFieldName = listResource.generateImageFieldName(field);
          const typeDefs = `
            type ${nodeType} implements Node {
              ${imageFieldName}: File @link(from: "fields.${imageFieldName}")
            }
          `
          createTypes(typeDefs);
        })
      }

    })
  })
};

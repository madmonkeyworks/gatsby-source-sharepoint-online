/**
 * @typedef {import('gatsby').SourceNodesArgs} Helpers
 * @typedef {import('@microsoft/microsoft-graph-client').Client} Client
 */

const generateListItemsUrl = (site, list, host) => {
  const base = `/sites/${host}`;
  return `${base}:/${site.relativePath}:/lists/${encodeURI(list.title)}/items`;
};
const fs = require("fs");
class Resource {
  /**
   * The type of this resource.
   * @type {'list' | 'drive'}
   */
  type;

  /**
   * Construct a resource.
   * @param {'list' | 'drive'} type The type of resource.
   */
  constructor(type) {
    this.type = type;

    if (type === "drive") {
      console.warn("Drives are not yet supported.");
    }
  }

  /**
   * Validates a resource item such as a list or drive definition.
   * @param {any} item The item to validate.
   */
  validate(item) {
    const isValid =
      this.type !== "drive" && item !== undefined && Boolean(item.title);
    if (!isValid) {
      console.warn(`Invalid resource item: ${JSON.stringify(item)}`);
    }

    return isValid;
  }

  /**
   * Create a graph resource request.
   * @param {string} host The SharePoint host.
   * @param {any} site The site definition object.
   * @param {Client} graph The graph client.
   * @param {Helpers} helpers The Gatsby sourceNode API helpers.
   */
  requestFactory(host, site, graph, helpers) {
    return async (item) => {
      let request = graph
        .api(generateListItemsUrl(site, item, host))
        .expand("fields");

      if (item.fields && Array.isArray(item.fields)) {
        request = request.expand(`fields($select=${item.fields.join(",")})`);
      }

      const normalizedListName = item.title.replace(" ", "");
      const normalizedCustomNodeName = item.customNodeName?.replace(" ", "") || null;

      /* 
      TODO!
      MS Graph paginates results by default. Max amount is 200 results per page. We need to pull all results at once. 
      Use PageIterator to iterate through all pages and pull all data.
      See https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/5183690123f81fc4170ce9e70cf0628d377f5fbc/docs/tasks/PageIterator.md
      */
      /*
      try {
        const entry = [];
        
        // Makes request to fetch mails list. Which is expected to have multiple pages of data.
        PageCollection = await request.get();
        
        // A callback function to be called for every item in the collection. This call back should return boolean indicating whether not to continue the iteration process.
        let callback = (data) => {
          entry.push(data);
          return true;
        };
        
        // Creating a new page iterator instance with client a graph client instance, page collection response from request and callback
        let pageIterator = new PageIterator(request, PageCollection, callback);
        
        // This iterates the collection until the nextLink is drained out.
        await pageIterator.iterate();
      } catch (e) {
        throw e;
      }
*/
      await request.get().then(entry => {
        const postNodeType = `SharePoint${normalizedCustomNodeName ? normalizedCustomNodeName : normalizedListName}List`;
        entry.value.forEach((data) => {
          // Create slug for the node
          if (item.createSlugs && Array.isArray(item.slugTemplate) && item.slugFieldName !== undefined) {
            // Construct slug from configuration
            const prepareSlug = [];
            item.slugTemplate.map( field => {
              prepareSlug.push(data.fields[field] ? data.fields[field] : field);
            });
            const slug = prepareSlug.join("-")
              .replace()
              .toLowerCase()
              .trim()
              .replace(/[^\w\s-]/g, "")
              .replace(/[\s_-]+/g, "-")
              .replace(/^-+|-+$/g, "");
            const slugFieldName = item.slugFieldName ? item.slugFieldName : "slug";
            // Add slug to node fields
            data.fields[slugFieldName] = slug;
            console.log(`Page for list ${item.title} created`) 
          } 
          helpers.actions.createNode({
            data,
            id: helpers.createNodeId(`${postNodeType}-${data.id}`),
            parent: null,
            children: [],
            internal: {
              type: postNodeType,
              content: JSON.stringify(data),
              contentDigest: helpers.createContentDigest(data),
            },
          });
        });
      })
    }
  }
}

exports.Resource = Resource;

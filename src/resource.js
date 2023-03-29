/**
 * @typedef {import('gatsby').SourceNodesArgs} Helpers
 * @typedef {import('@microsoft/microsoft-graph-client').Client} Client
 */

/**
 * Generate a list URL used for requests 
 * @param {any} site The item to validate.
 * @param {string} list The item to validate.
 * @param {string} host The item to validate.
*/
const generateListItemsUrl = (site, list, host) => {
  const base = `/sites/${host}`;
  return `${base}:/${site.relativePath}:/lists/${encodeURI(list.title)}/items`;
};

/**
 * Generate a URL used for requesting list's site asset drive item (e.g. image)
 * @param {any} site The item to validate.
 * @param {string} itemID The item to validate.
 * @param {string} host The item to validate.
*/
const generateListAssetItemUrl = (site, itemID, host) => {
  const base = `/sites/${host}`;
  return `${base}:/${site.relativePath}:/lists/${encodeURI("Site Assets")}/items/${itemID}/driveItem/content`;
};

class Resource {
  /**
   * The type of this resource.
   * @type { 'list' }
   */
  type;

  /**
   * Construct a resource.
   * @param { 'list' } type The type of resource.
   */
  constructor(type) {
    this.type = type;
    this.imageFieldSuffix = 'Image';
  }

  /**
   * Validates a list resource item 
   * @param {any} list The item to validate.
   */
  validateList(list) {
    const isValid =
      this.type === "list" && list !== undefined && Boolean(list.title);
    if (!isValid) {
      console.warn(`Invalid resource list: ${JSON.stringify(list)}`);
    }

    return isValid;
  }

  /**
   * Normalize a list name 
   * @param {Object} list The list to retrieve title from and normalize it.
   */
  normalizeListName(list) {
    let name = list.customNodeName || list.title
    return name.replace(" ", "");
  }
  
  /**
   * Generate a list node type name
   * @param {string} listName The list to normalize.
   */
  generateListNodeType(listName) {
    return `SharePoint${listName}List`
  }
  /**
  * Generate image field name for custom schema
  * @param {string} field The item to validate.
  */
  generateImageFieldName = (field) => {
    return `${field}${this.imageFieldSuffix}`;
  }

  /**
   * LISTS RESOURCE CLASS
   * Create a graph resource request.
   * @param {string} host The SharePoint host.
   * @param {any} site The site definition object.
   * @param {Client} graph The graph client.
   * @param {Helpers} helpers The Gatsby sourceNode API helpers.
   */
  requestLists(host, site, graph, helpers) {
    
    return async (list) => {
      
      const listName = this.normalizeListName(list)
      
      let request = graph
        .api(generateListItemsUrl(site, list, host))

      // Select fields
      if (list.fields && Array.isArray(list.fields)) {
        const selects = [];
        list.fields.forEach(f => {
          if (typeof f === 'string' ) {
            selects.push(f)
          }
          if (typeof f === 'object' ) {
            selects.push(f.fieldName)
          }
        })
        if (selects) {
          request = request.expand(`fields($select=${selects.join(",")})`);
        } else {
          request = request.expand("fields");
        }
      }
      
      await request.get().then(entry => {
        const nodeType = this.generateListNodeType(listName);
        entry.value.forEach((data) => {
          
          const nodeId = helpers.createNodeId(`${nodeType}-${data.id}`);
          const node = {
            data,
            id: nodeId,
            parent: null,
            children: [],
            internal: {
              type: nodeType,
              content: JSON.stringify(data),
              contentDigest: helpers.createContentDigest(data),
            },
          };
          helpers.actions.createNode(node);

          // Add image nodes for explicitly defined fields
          const imageFields = list.fields.filter(f => f?.fieldType === "image").map(f => f.fieldName);
          imageFields.forEach(field => {
            if (!data.fields[field]) {
              return;
            }
            try {
              const { id } = JSON.parse(data.fields[field]);
              const url = generateListAssetItemUrl(site, id, host)
              this.createImageNodeField(url, field, node, graph, helpers)();
            } catch (err) {
              console.error(`Couldn't retrieve the image for field "${field}", list "${list.title}", site "${site.name}"`, err);
            }
          })
        });
      });
    };
  }

  /**
   * Create a graph file resource request
   * @param {string} url The SharePoint host.
   * @param {string} fieldName The field name.
   * @param {Node} node The parent node.
   * @param {Client} graph The graph client.
   * @param {Helpers} helpers The Gatsby sourceNode API helpers.
   */
  createImageNodeField(url, fieldName, node, graph, helpers) {
    return async () => {
      const { createFileNodeFromBuffer } = require("gatsby-source-filesystem");
      const { createNode, createNodeField } = helpers.actions;
      const { createNodeId, getCache } = helpers;

      const attachField = async (arrayBuffer) => {
        const fileNode = await createFileNodeFromBuffer({
          buffer: arrayBuffer,
          getCache,
          createNode,
          createNodeId,
          parentNodeId: node.id
        })
        // if the file was created, extend the node with "File"
        if (fileNode) {
          createNodeField({ node, name: `${fieldName}${this.imageFieldSuffix}`, value: fileNode.id })
        } 
      }

      await graph.api(url)
        .getStream()
        .then((stream) => {
          let chunks = [];
          stream.on('data', chunk => {
              chunks.push(chunk);
          });
          stream.on('end', () => {
              const buffer = Buffer.concat(chunks);
              attachField(buffer)
          });
        })
        .catch((err) => {
          console.log(err);
        });
    }
  }
}

exports.Resource = Resource;

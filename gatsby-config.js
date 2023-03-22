// loads environment variables into NodeJS
require("dotenv").config({
  path: `.env.${process.env.NODE_ENV}`,
})

module.exports = {
  siteMetadata: {
    title: `COFRA Group Vacancies Portal`,
    slogan: `Seeking opportunities`, 
    description: `This online meeting place is designed for COFRA Group employees to browse through - and get inspired by - the various available role opportunities within our group of companies. Start now to broaden your horizon.`
  },
  plugins: [
    {
      resolve: "gatsby-source-sharepoint-online",
      options: {
        host: process.env.SHAREPOINT_HOST,
        appId: process.env.APP_ID,
        appSecret: process.env.APP_SECRET,
        appRedirectUri: process.env.APP_REDIRECT_URI,
        tenantId: process.env.TENANT_ID,
        sites: [
          {
            name: process.env.SHAREPOINT_SITE_NAME,
            relativePath: process.env.SHAREPOINT_SITE_RELATIVE_PATH,
            lists: [
              {
                title: 'Pages',
                fields: ['Title', 'Pagetitle', 'Subtitle', 'Slug', 'Description', 'Parent']
              },
              {
                title: 'Menus',
                fields: ['Location', 'Title', 'Page', 'Page_x003a__x0020_Slug', 'Order', 'Parent']
              },
              {
                title: 'Variables',
                fields: ['Title', 'Value']
              }
              // {
              //   title: "Vacancies",
              //   customNodeName: "Vacancies",
              //   fields: [],
              //   createSlugs: true,
              //   slugFieldName: "slug",
              //   slugTemplate: ["Position", "Title"],
              // }
            ],
          },
        ],
      },
    },
  ],
};

FROM m365pnp/spfx:latest
COPY package*.json ./
RUN npm install gulp-cli --global
RUN npm install yo --global
RUN npm install @microsoft/generator-sharepoint --global
RUN npm install
COPY . .
EXPOSE 4321
CMD [ "gulp", "serve" ]
#docker run -it --rm --name spfx -v ./:/usr/app/spfx -p 4321:4321 -p 35729:35729 m365pnp/spfx
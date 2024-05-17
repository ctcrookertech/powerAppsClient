export default class PowerAppsClient {
    constructor() {
        this.batchSize = 10;
        this.connectors = [];
        this.connectorDetails = [];
        this.powerAppsApiRequestBaseUrl = 'https://api.powerapps.com/providers/Microsoft.PowerApps/apis/';
        this.powerAppsApiRequestParameters = '?api-version=2023-06-01&$filter=environment+eq+%27~Default%27';
    }

    getConnectors() {
        return structuredClone(this.connectors);
    }

    getConnectorDetails() {
        return structuredClone(this.connectorDetails);
    }

    getConnectorUrl(id) {
        return `${this.powerAppsApiRequestBaseUrl}${id}${this.powerAppsApiRequestParameters}`;
    }

    getInterfaces(interfaces) {
        const modes = [];

        if (interfaces) {
            for (const property in interfaces) {
                if (!JSON.stringify(interfaces[property].revisions).match(/"deprecated":true/)) {
                    if (property.match(/blob/i)) {
                        modes.push('blob');
                    } else if (property.match(/tabular/i)) {
                        modes.push('tabular');
                    } else {
                        modes.push(property);
                    }
                }
            }
        }

        return modes.sort().join(', ');
    }

    addConnectorDetails(headers) {
        for (var count = 0; count < this.batchSize && this.connectors.length > 0; count++) {
            const connector = this.connectors.pop();
            fetch(this.getConnectorUrl(connector.id), headers)
            .then((response) => response.json())
            .then((data) => {
                console.log(data.name);
                const properties = data.properties;
                const swagger = properties.swagger;
                const metadata = swagger['x-ms-connector-metadata'];
                var categories = '';
                var website = '';
                if (metadata) {
                    metadata.forEach((property) => {
                        if (property.propertyName.match(/categories/i)) {
                            categories = property.propertyValue.split(';').sort().join(', ');
                        } else if (property.propertyName.match(/website/i)) {
                            website = property.propertyValue;
                        }
                    });
                }
                connector.categories = categories;
                connector.website = website;
                const operations = [];
                for (var path in swagger.paths) {
                    const operationPath = swagger.paths[path];
                    for (var method in operationPath) {
                        const operation = operationPath[method];
                        operations.push(`${operation.operationId}: ${operation.description}`);
                    }
                }
                connector.operations = operations.join('\n');
                this.connectorDetails.push(connector);
            });
        }

        if (this.connectors.length > 0) {
            setTimeout(() => this.addConnectorDetails(headers), 100);
        }
    }

    fetchConnectors(token) {
        fetch(this.getConnectorUrl(''), { headers: { Authorization: `Bearer ${token}` } })
        .then((response) => response.json())
        .then((data) => {
            data.value.forEach((connector) => {
                if (!connector.isCustomApi) {
                    const properties = connector.properties;
                    this.connectors.push({
                        id: connector.name,
                        name: properties.displayName,
                        description: properties.description,
                        interfaces: this.getInterfaces(properties.interfaces),
                        capabilities: properties.capabilities.sort().join(', '),
                        allowed: (!properties.scopes ? '' : properties.scopes.will.join('\n')),
                        denied: (!properties.scopes ? '' : properties.scopes.wont.join('\n'))
                    });
                }
            });

            const headers = { headers: { Authorization: `Bearer ${token}` } };
            this.addConnectorDetails(headers);
        });
    }
}
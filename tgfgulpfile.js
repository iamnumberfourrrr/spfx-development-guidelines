'use strict';
const fs = require('fs');
const uuid = require('uuid');
const gulp = require('gulp');
const path = require('path');
const build = require('@microsoft/sp-build-web');
const gulpUtil = require('gulp-util');
var ncp = require('ncp').ncp;
var shell = require('shelljs');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
const merge = require('webpack-merge');
const TerserPlugin = require('terser-webpack-plugin-legacy');

// Retrieve the current build config and check if there is a `warnoff` flag set
const crntConfig = build.getConfig();
const warningLevel = crntConfig.args["warnoff"];

// Extend the SPFx build rig, and overwrite the `shouldWarningsFailBuild` property
if (warningLevel) {
    class CustomSPWebBuildRig extends build.SPWebBuildRig {
        setupSharedConfig() {
            build.log("IMPORTANT: Warnings will not fail the build.")
            build.mergeConfig({
                shouldWarningsFailBuild: false
            });
            super.setupSharedConfig();
        }
    }

    build.rig = new CustomSPWebBuildRig();
}

build.configureWebpack.setConfig({
    additionalConfiguration: function (config) {
        let newConfig = config;
        config.plugins.forEach((plugin, i) => {
            if (plugin.options && plugin.options.mangle) {
                config.plugins.splice(i, 1);
                newConfig = merge(config, {
                    plugins: [
                        new TerserPlugin()
                    ]
                });
            }
        });

        return newConfig;
    }
});

build.initialize(gulp);

function updateProvisioningTemplate(originComponentDict, targetComponentDict, file) 
{
    //try {
        fs.accessSync(file, fs.constants.W_OK);
		/*
    } catch 
	{
        setTimeout(() => updateProvisioningTemplate(originComponentDict, targetComponentDict, file), 1000);
        return;
    }
	*/
    console.log('Patching provisioning template ' + file + '...')
    const keys = Object.keys(originComponentDict);

    let result = fs.readFileSync(file, 'utf8');
    for (let key of keys) {
        if (result.indexOf(originComponentDict[key]) >= 0) {
            console.log('Found usage of component ' + key + '. Value in origin component:' + originComponentDict[key] + ' --- value in target component:' + targetComponentDict[key]);
        }
        result = result.replace(originComponentDict[key], targetComponentDict[key]);
        while (result.indexOf(originComponentDict[key]) >= 0) {
            result = result.replace(originComponentDict[key], targetComponentDict[key]);
        }
    }
    fs.writeFileSync(file, result, 'utf8');
    console.log('Patching provisioning template ' + file + ' OK');
}

function copyManifest(targetEnv, updateId = true) {
    const compDict = {};
    const configPath = './config/config.json';
    const packageCfgPath = './config/package-solution.json';

    const envDir = './env-build-config/' + targetEnv;
    const envPackageSlnPath = envDir + '/config/package-solution.json';
    const dir = path.dirname(envPackageSlnPath);
    if (!fs.existsSync(dir)) {
        shell.mkdir('-p', dir);
    }

    if (!fs.existsSync(envPackageSlnPath)) {
        const packageContent = require(packageCfgPath);
        if (updateId) {
            packageContent.solution.id = uuid.v4();
            packageContent.solution.name += '-' + targetEnv.toLowerCase();
            packageContent.solution.features[0].title = targetEnv + " - " + packageContent.solution.features[0].title;
            packageContent.solution.features[0].id = uuid.v4();
            packageContent.paths.zippedPackage = 'solution/tgf-web-parts.' + targetEnv + '.sppkg';
            fs.writeFileSync(envPackageSlnPath, JSON.stringify(packageContent, null, 4), 'utf8');
        } else {
            fs.copyFileSync(packageCfgPath, envPackageSlnPath);
        }
    }

    if (!fs.existsSync(envDir + '\\assets')) {
        shell.mkdir('-p', envDir + '\\assets');
    }

    fs.copyFileSync('.\\sharepoint\\assets\\elements.xml', envDir + '\\assets\\elements.xml');
    fs.copyFileSync('.\\sharepoint\\assets\\ClientSideInstance.xml', envDir + '\\assets\\ClientSideInstance.xml');

    var config = require(configPath);
    for (var key of Object.keys(config.bundles)) {
        const components = config.bundles[key].components;
        for (var c of components) {
            var content = require(c.manifest);
            var targetManifest = c.manifest;
            targetManifest = envDir + targetManifest.substr(1);
            const dir = path.dirname(targetManifest);
            if (!fs.existsSync(dir)) {
                shell.mkdir('-p', dir);
            }

            if (!fs.existsSync(targetManifest)) {
                if (updateId) {
                    content.id = uuid.v4();
                    if (content.preconfiguredEntries) {
                        content.preconfiguredEntries[0].title.default = `${targetEnv} - ${content.preconfiguredEntries[0].title.default}`
                    }
                    fs.writeFileSync(targetManifest, JSON.stringify(content, null, 4), 'utf8');
                } else {
                    fs.copyFileSync(c.manifest, targetManifest);
                }
            } else {
                content = require(targetManifest)
            }
            compDict[content.alias] = content.id;
        }
    }

    fs.writeFileSync(envDir + '\\components.json', JSON.stringify(compDict, null, 4), 'utf-8');
}

gulp.task('tgf-package-env', function () {
    var targetBuildEnv = (process.argv[5] || 'DEV').trim().toUpperCase();
    console.log('Target Environment: ', targetBuildEnv);
    copyManifest("Origin", false);
    copyManifest(targetBuildEnv, true);
    const envDir = '.\\env-build-config\\' + targetBuildEnv;
    const originComponentDict = JSON.parse(fs.readFileSync('.\\env-build-config\\Origin\\components.json', 'utf8'));
    const targetComponentDict = JSON.parse(fs.readFileSync(envDir + '\\components.json', 'utf8'));
    updateProvisioningTemplate(originComponentDict, targetComponentDict, envDir + '\\assets\\elements.xml');
    updateProvisioningTemplate(originComponentDict, targetComponentDict, envDir + '\\assets\\ClientSideInstance.xml');

    const targetFolder = './';
    const srcFolder = './env-build-config/' + targetBuildEnv;
    ncp(envDir + '\\assets', '.\\sharepoint\\assets');
    let fc = fs.readFileSync('.\\src\\common\\utils.tsx', 'utf8');
    fc = fc.replace(/const env = "(\w+)";/, `const env = "${targetBuildEnv}";`);
    fs.writeFileSync('.\\src\\common\\utils.tsx', fc, 'utf8');
    return ncp(srcFolder, targetFolder, err => {
        console.log('Configuration for environment ', targetBuildEnv, ' ready. Run Bundle and Package-solution to create package');
    });
});

gulp.task('tgf-update-provisioning-template-env', function () {
    var targetBuildEnv = (process.argv[5] || 'DEV').trim().toUpperCase();
    const provisioningRootDir = '..\\ProvisioningTemplates';
    const envProvisionDir = provisioningRootDir + '\\' + targetBuildEnv;
    console.log('Target Environment: ', targetBuildEnv);
    if (targetBuildEnv !== "DEV") {
        if (!fs.existsSync(envProvisionDir)) {
            shell.mkdir('-p', envProvisionDir);
        }
        fs.copyFileSync(provisioningRootDir + '\\ckms-portal-sitecollection.xml', envProvisionDir + '\\ckms-portal-sitecollection.xml');
        fs.copyFileSync(provisioningRootDir + '\\config-sitecollection.xml', envProvisionDir + '\\config-sitecollection.xml');
		fs.copyFileSync(provisioningRootDir + '\\searchsitecollectionTEST.xml', envProvisionDir + '\\searchsitecollectionTEST.xml');
        const targetComponentDict = JSON.parse(fs.readFileSync('.\\env-build-config\\' + targetBuildEnv + '\\components.json', 'utf8'));
        const originComponentDict = JSON.parse(fs.readFileSync('.\\env-build-config\\Origin\\components.json', 'utf8'));
        updateProvisioningTemplate(originComponentDict, targetComponentDict, envProvisionDir + '\\ckms-portal-sitecollection.xml');        
        ncp(provisioningRootDir + '\\Resource', envProvisionDir + '\\Resource', () => {
            updateProvisioningTemplate(originComponentDict, targetComponentDict, envProvisionDir + '\\Resource\\CaseSiteTemplate\\case-sitecollection_3.15.1911.0.xml');
			updateProvisioningTemplate(originComponentDict, targetComponentDict, envProvisionDir + '\\Resource\\CaseSiteTemplate\\audit-sitecollection_3.15.1911.0.xml');
        });
    }
});

gulp.task('tgf-package-env-reset', function () {
    const targetFolder = './';
    const orgFolder = './env-build-config/Origin';
    ncp(orgFolder + '/assets', '.\\sharepoint\\assets');
    let fc = fs.readFileSync('.\\src\\common\\utils.tsx', 'utf8');
    fc = fc.replace(/const env = "(\w+)";/, 'const env = "DEV";');
    fs.writeFileSync('.\\src\\common\\utils.tsx', fc, 'utf8');
    return ncp(orgFolder, targetFolder, err => {
        console.log('Configuration resetted');
    });
});
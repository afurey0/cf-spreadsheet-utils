/**
 * ColdFusion implementation for the POI library's Temp File Creation Strategy.
 * Can be instantiated as a Java object using ColdFusion's built-in createDynamicProxy function like so: {@code createDynamicProxy("cfc.ColdFusionTempFileCreationStrategy", ["org.apache.poi.util.TempFileCreationStrategy"])}
 * @see {@link https://poi.apache.org/apidocs/dev/org/apache/poi/util/TempFileCreationStrategy.html}
 */
component {

	public any function createTempDirectory(required string prefix) {
		local.directory = makeDirectory(getRoot(), arguments.prefix & "_" & lCase(createUUID()));
		local.directory.deleteOnExit();
		return local.directory;
	}

	public any function createTempFile(required string prefix, required string suffix) {
		local.file = createObject("java", "java.io.File").createTempFile(arguments.prefix, arguments.suffix, getRoot());
		local.file.deleteOnExit();
		return local.file;
	}

	private any function getRoot() {
		return makeDirectory(createObject("java", "java.io.File").init(getTempDirectory()), "poifiles");
	}

	private any function makeDirectory(required any base, required string name) {
		local.directory = createObject("java", "java.io.File").init(arguments.base, arguments.name);
		local.exists = local.directory.exists() or local.directory.mkdirs();
		if (not local.exists) {
			throw("FileException", "Could not create temporary directory.", local.base);
		} else if (not local.directory.isDirectory()) {
			throw("FileException", "Path exists but is not a directory.", local.base);
		}
		return local.directory;
	}

}
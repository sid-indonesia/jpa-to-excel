package org.sidindonesia.jpatoexcel.util;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.stream.Collectors;

import javax.persistence.Entity;

import org.reflections.Reflections;
import org.reflections.scanners.SubTypesScanner;
import org.reflections.scanners.TypeAnnotationsScanner;
import org.springframework.stereotype.Repository;

public final class ReflectionsUtil {

	private ReflectionsUtil() {
		// Utility class
	}

	private static final TypeAnnotationsScanner TYPE_ANNOTATIONS_SCANNER = new TypeAnnotationsScanner();
	private static final SubTypesScanner SUB_TYPES_SCANNER = new SubTypesScanner();

	private static Set<Class<?>> repositoryClasses;

	private static Set<Class<?>> entityClasses;

	public static Set<Class<?>> getAllEntityClasses(String jpaEntityPackageName) {
		if (entityClasses == null) {
			entityClasses = new Reflections(jpaEntityPackageName, TYPE_ANNOTATIONS_SCANNER, SUB_TYPES_SCANNER)
				.getTypesAnnotatedWith(Entity.class).stream()
				.sorted((c1, c2) -> c1.getSimpleName().compareTo(c2.getSimpleName()))
				.collect(Collectors.toCollection(LinkedHashSet::new));
		}

		return entityClasses;
	}

	public static Class<?> getRepositoryClassOfEntity(Class<?> entityClass, String jpaRepositoryPackageName) {
		if (repositoryClasses == null) {
			repositoryClasses = new Reflections(jpaRepositoryPackageName, TYPE_ANNOTATIONS_SCANNER, SUB_TYPES_SCANNER)
				.getTypesAnnotatedWith(Repository.class);
		}

		return repositoryClasses.stream().filter(repositoryClass -> {
			Type type = ((ParameterizedType) repositoryClass.getGenericInterfaces()[0]).getActualTypeArguments()[0];
			return entityClass.getTypeName().equals(type.getTypeName());
		}).findAny().orElseThrow();
	}
}

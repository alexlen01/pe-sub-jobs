FROM maven:3.9-eclipse-temurin-25 AS build
WORKDIR /app
COPY pom.xml ./
RUN mvn dependency:go-offline -q
COPY src ./src
RUN mvn -q package -DskipTests

FROM eclipse-temurin:25-jre
WORKDIR /app
COPY --from=build /app/target/pe-sub-jobs-1.0.0.jar app.jar
EXPOSE 3003
ENTRYPOINT ["java", "-jar", "app.jar"]

FROM maven:3.9-eclipse-temurin-25 AS build
WORKDIR /app
COPY pom.xml ./
RUN mvn dependency:go-offline -q
COPY src ./src
COPY data ./data
RUN mvn -q package -DskipTests

FROM eclipse-temurin:25-jre
WORKDIR /app
COPY --from=build /app/target/pe-sub-jobs-1.0.0.jar app.jar
COPY --from=build /app/data ./data
EXPOSE 3003
# Size the heap from the container memory limit (Temurin default is only 25%).
ENTRYPOINT ["java", "-XX:MaxRAMPercentage=75.0", "-jar", "app.jar"]

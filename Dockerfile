FROM php:8.2-cli-alpine

# Install system dependencies and required PHP extensions
RUN apk add --no-cache \
    sqlite-libs \
    libzip \
    libpng \
    libjpeg-turbo \
    freetype \
    icu-libs \
    libxml2 \
    curl \
    git \
    unzip \
    && apk add --no-cache --virtual .build-deps \
    $PHPIZE_DEPS \
    sqlite-dev \
    libzip-dev \
    libpng-dev \
    libjpeg-turbo-dev \
    freetype-dev \
    icu-dev \
    libxml2-dev \
    && docker-php-ext-configure gd --with-freetype --with-jpeg \
    && docker-php-ext-install -j$(nproc) \
    pdo_sqlite \
    zip \
    gd \
    intl \
    bcmath \
    xml \
    && apk del .build-deps

# Install Composer
COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

WORKDIR /app

# Copy Composer files first for layer caching
COPY composer.json composer.lock ./

# Install PHP dependencies without dev packages
RUN composer install --no-dev --optimize-autoloader --no-interaction

# Copy project source files
COPY . .

# Create application data directories
RUN mkdir -p data downloads logs temp uploads \
    && chmod -R 777 data downloads logs temp uploads

# Run installation and start polling daemon
CMD ["sh", "-c", "php install.php && php bot.php"]
